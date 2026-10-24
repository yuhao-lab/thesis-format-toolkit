# -*- coding: utf-8 -*-
"""
论文格式一键排版

命令行用法：
    python format_docx.py 我的论文.docx
    python format_docx.py 我的论文.docx -o 排好版.docx -r rules.yaml

图形界面用法：
    python gui.py
"""

import argparse
from collections import Counter
from pathlib import Path

import yaml
import re

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Cm, Pt

import normalize as nz

# 中文号数 → 磅值，rules.yaml 里也可以直接写数字
CN_SIZE = {
    "初号": 42, "小初": 36, "一号": 26, "小一": 24,
    "二号": 22, "小二": 18, "三号": 16, "小三": 15,
    "四号": 14, "小四": 12, "五号": 10.5, "小五": 9,
}

ALIGN = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

# rules.yaml 里哪一段对应 Word 里的哪个样式名
# 注意：一个 Word 样式只能被配置一次，重复映射会互相覆盖
STYLE_MAP = {
    "title": ["Title", "标题"],
    "body": ["Normal", "正文"],
    "h1": ["Heading 1", "标题 1"],
    "h2": ["Heading 2", "标题 2"],
    "h3": ["Heading 3", "标题 3"],
    "caption": ["Caption", "题注"],
    "reference": ["Bibliography", "参考文献"],   # 文档里没有这个样式就跳过
}


def to_pt(value):
    """把 12 / '小四' / '12pt' 统一转成磅值"""
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace("pt", "").replace("磅", "")
    if text in CN_SIZE:
        return CN_SIZE[text]
    return float(text)


def set_line_spacing(pf, value):
    """小于等于 5 当作倍数，大于 5 当作固定磅值"""
    if isinstance(value, (int, float)) and value <= 5:
        pf.line_spacing = float(value)
    else:
        pf.line_spacing = Pt(to_pt(value))


def set_style_format(style, cfg):
    """把一份配置应用到某个 Word 样式上"""
    # ---------- 字体 ----------
    font_cn = cfg.get("font_cn", "宋体")
    font_en = cfg.get("font_en", "Times New Roman")
    style.font.name = font_en                      # 西文
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn("w:eastAsia"), font_cn)          # 中文（必须单独设）
    rfonts.set(qn("w:ascii"), font_en)
    rfonts.set(qn("w:hAnsi"), font_en)

    # ---------- 字号与粗细 ----------
    style.font.size = Pt(to_pt(cfg.get("size", 12)))
    style.font.bold = bool(cfg.get("bold", False))

    # ---------- 段落 ----------
    pf = style.paragraph_format
    pf.alignment = ALIGN.get(str(cfg.get("align", "left")).lower(),
                             WD_ALIGN_PARAGRAPH.LEFT)
    set_line_spacing(pf, cfg.get("line_spacing", 1.5))
    pf.space_before = Pt(float(cfg.get("space_before", 0)))
    pf.space_after = Pt(float(cfg.get("space_after", 0)))

    # ---------- 首行缩进（按字符，Word 原生支持） ----------
    chars = float(cfg.get("first_line_chars", 0))
    ppr = style.element.get_or_add_pPr()
    ind = ppr.get_or_add_ind()
    size_pt = to_pt(cfg.get("size", 12))
    if chars > 0:
        ind.set(qn("w:firstLineChars"), str(int(chars * 100)))
        ind.set(qn("w:firstLine"), str(int(chars * size_pt * 20)))
    else:
        ind.set(qn("w:firstLineChars"), "0")
        ind.set(qn("w:firstLine"), "0")


def add_page_number(section, align="center", size=9):
    """在页脚插入 PAGE 域。页脚已有文字时不覆盖，只记一笔。"""
    footer = section.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()

    if p.text.strip():
        return f"页脚已有内容「{p.text.strip()[:16]}」，跳过页码"

    p.alignment = ALIGN.get(align, WD_ALIGN_PARAGRAPH.CENTER)

    for run in list(p.runs):
        run._r.getparent().remove(run._r)

    run = p.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    placeholder = OxmlElement("w:t")
    placeholder.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")

    for node in (begin, instr, sep, placeholder, end):
        run._r.append(node)

    # 页码字号：很多学校要求小五号
    try:
        run.font.size = Pt(float(size))
    except Exception:
        pass
    return None


def ensure_default_style(doc):
    """有些文档（从业务系统导出的居多）styles.xml 里一个样式都没定义，
    段落全靠直接格式撑着。这种情况下改样式表改不动任何东西。
    这里补一个 Normal 并设为默认，后面改它就能影响全文。

    返回 True 表示补过了。
    """
    try:
        doc.styles["Normal"]
        return False
    except KeyError:
        pass

    style = doc.styles.add_style("Normal", WD_STYLE_TYPE.PARAGRAPH)
    style.element.set(qn("w:default"), "1")
    return True


def apply_rules(doc, rules):
    """把规则应用到已打开的文档对象，返回统计结果。"""
    notes = []

    # ---------- 0. 先数坏引用，再补样式表 ----------
    broken = sum(1 for p in doc.paragraphs if p.style is None)
    repaired = ensure_default_style(doc)
    if repaired:
        notes.append("这份文档没有样式定义（多来自业务系统导出），已自动补上默认样式")
    if broken:
        notes.append(f"{broken} 段原本引用了不存在的样式，已按正文处理")

    # ---------- 1. 页面 ----------
    page = rules.get("page", {})
    for section in doc.sections:
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(float(page.get("margin_top", 2.5)))
        section.bottom_margin = Cm(float(page.get("margin_bottom", 2.5)))
        section.left_margin = Cm(float(page.get("margin_left", 3.0)))
        section.right_margin = Cm(float(page.get("margin_right", 2.5)))
        section.gutter = Cm(float(page.get("gutter", 0)))
        if page.get("page_number", True):
            msg = add_page_number(section, page.get("page_number_align", "center"),
                                  page.get("page_number_size", 9))
            if msg:
                notes.append(msg)

    # ---------- 2. 样式 ----------
    applied, missing = [], []
    for key, names in STYLE_MAP.items():
        if key not in rules:
            continue
        cfg = rules[key]
        target = None
        for name in names:
            try:
                target = doc.styles[name]
                break
            except KeyError:
                continue
        if target is None:
            missing.append(key)
            continue
        set_style_format(target, cfg)
        applied.append(target.name)

    # ---------- 2.5 网页 / 列表类正文样式 ----------
    # 从网页、公众号、网页版 AI 复制过来的段落，常带 Normal (Web)、Body Text、
    # List Paragraph 这些样式。它们不在 STYLE_MAP 里就永远改不到，
    # 结果是同一篇文档里两套正文格式，观感割裂。
    BODY_LIKE = [
        "Normal (Web)", "Body Text", "Body Text 2", "Body Text 3",
        "List Paragraph", "No Spacing",
        "正文文本", "列表段落", "无间隔",
    ]
    body_cfg = rules.get("body")
    if body_cfg:
        for name in BODY_LIKE:
            try:
                target = doc.styles[name]
            except KeyError:
                continue
            set_style_format(target, body_cfg)
            applied.append(target.name)

    # ---------- 3. 表格内容 ----------
    # 表格单元格会继承正文的「首行缩进 2 字符」，代码块（1x1 表格）还会被套成宋体，
    # 这两样都得单独处理，否则表格里的字往里缩、代码变成比例字体。
    code_tables = 0
    for table in doc.tables:
        is_code = len(table.rows) == 1 and len(table.columns) == 1
        if is_code:
            code_tables += 1
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    pf = para.paragraph_format
                    pf.first_line_indent = Pt(0)
                    pf.left_indent = Pt(0)
                    for run in para.runs:
                        run.font.size = Pt(9.5 if is_code else 10.5)
                        if is_code:
                            run.font.name = "Consolas"
                            rpr = run._r.get_or_add_rPr()
                            rpr.get_or_add_rFonts().set(qn("w:eastAsia"), "宋体")
        if not is_code and len(table.rows):
            for cell in table.rows[0].cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.bold = True

    # ---------- 3.5 摘要 / 关键词标签 ----------
    abstract_count = format_abstract_labels(doc, rules.get("abstract_label", {}))

    # ---------- 4. 统计 ----------
    # 补过默认样式后，原本 p.style 为 None 的段落会回退到 Normal，这里再兜一层。
    counts = Counter()
    weird = 0
    for p in doc.paragraphs:
        st = p.style
        if st is None:
            weird += 1
            counts["__none__"] += 1
        else:
            counts[st.name] += 1

    h1 = counts.get("Heading 1", 0) + counts.get("标题 1", 0)
    h2 = counts.get("Heading 2", 0) + counts.get("标题 2", 0)
    h3 = counts.get("Heading 3", 0) + counts.get("标题 3", 0)
    body = counts.get("Normal", 0) + counts.get("正文", 0)

    # 直接格式占比：run 自己带字号的，说明是逐段手动设的格式，
    # 这种情况下改样式表压不住，正文格式改不动。
    runs = [r for p in doc.paragraphs for r in p.runs]
    direct = sum(1 for r in runs if r.font.size is not None)
    direct_ratio = (direct / len(runs)) if runs else 0.0

    return {
        "applied": list(dict.fromkeys(applied)),
        "missing": missing,
        "h1": h1,
        "h2": h2,
        "h3": h3,
        "body": body,
        "weird": weird,
        "direct_ratio": direct_ratio,
        "code_tables": code_tables,
        "abstract_labels": abstract_count,
        "notes": notes,
    }


ABSTRACT_LABEL = re.compile(r"^(摘\s*要|关\s*键\s*词|Abstract|Key\s*words)\s*[:：]\s*", re.I)


def format_abstract_labels(doc, cfg=None):
    """把「摘要：」「关键词：」这几个标签设成黑体、顶格，正文内容保持宋体。

    学校要求标签顶格排，但整段又属于正文，只能拆 run。
    """
    cfg = cfg or {}
    label_font = cfg.get("font_cn") or "黑体"      # 默认黑体，别跟着正文走
    run_size = float(cfg.get("size", 12))
    n = 0

    for p in doc.paragraphs:
        text = p.text
        m = ABSTRACT_LABEL.match(text)
        if not m:
            continue

        head = text[:m.end()]
        tail = text[m.end():]

        if not p.runs:
            continue
        for r in p.runs[1:]:
            r._r.getparent().remove(r._r)
        p.runs[0].text = head
        r1 = p.runs[0]
        r2 = p.add_run(tail) if tail else None

        # 标签：黑体
        r1.font.size = Pt(run_size)
        rpr = r1._r.get_or_add_rPr()
        rfonts = rpr.get_or_add_rFonts()
        rfonts.set(qn("w:eastAsia"), label_font)
        rfonts.set(qn("w:ascii"), "Times New Roman")
        rfonts.set(qn("w:hAnsi"), "Times New Roman")
        r1.bold = False

        # 内容：宋体
        if r2 is not None:
            r2.font.size = Pt(run_size)
            rpr2 = r2._r.get_or_add_rPr()
            rf2 = rpr2.get_or_add_rFonts()
            rf2.set(qn("w:eastAsia"), "宋体")
            rf2.set(qn("w:ascii"), "Times New Roman")
            rf2.set(qn("w:hAnsi"), "Times New Roman")

        # 顶格
        p.paragraph_format.first_line_indent = Pt(0)
        ppr = p._p.get_or_add_pPr()
        ind = ppr.get_or_add_ind()
        ind.set(qn("w:firstLineChars"), "0")
        ind.set(qn("w:firstLine"), "0")
        n += 1

    return n


def should_normalize(doc, mode="auto"):
    """判断要不要先重建文档结构。

    auto：直接格式占比过半就重建（这类文档改样式表没用）
    on  ：强制重建
    off ：不重建，只改样式表
    """
    mode = str(mode).lower()
    if mode == "on":
        return True
    if mode == "off":
        return False

    try:
        doc.styles["Normal"]
    except KeyError:
        return True          # 样式表都没了，必须重建

    runs = [r for p in doc.paragraphs for r in p.runs]
    if not runs:
        return False
    direct = sum(1 for r in runs if r.font.size is not None)
    if (direct / len(runs)) > 0.5:
        return True

    # 标题样式只做了一半的文档（「一、」有的挂了标题、有的还是正文），
    # 直接格式占比不高，但结构是殘的，同样要重建。
    return nz.count_unstyled_numbered(doc) >= 1


def process(src, out, rules):
    """读文件 → 必要时重建结构 → 应用规则 → 保存。"""
    src, out = Path(src), Path(out)
    if not src.exists():
        raise FileNotFoundError(f"找不到文件：{src}")
    if src.suffix.lower() != ".docx":
        raise ValueError("只支持 .docx 文件。如果是 .doc，先在 Word 里另存为 .docx。")

    doc = Document(str(src))

    report = None
    if should_normalize(doc, rules.get("normalize", "auto")):
        report = nz.normalize(doc)

    stats = apply_rules(doc, rules)
    stats["normalize"] = report
    doc.save(str(out))
    return stats


def load_rules(path):
    return yaml.safe_load(Path(path).read_text(encoding="utf-8"))


def _safe_stdout():
    """打包成无控制台的 exe 时，sys.stdout 可能为 None；
    有控制台时又可能是 GBK，打不出 ⚠ 之类的字符。这里统一兜住。"""
    import os
    import sys

    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w", encoding="utf-8")
    else:
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def main():
    _safe_stdout()

    ap = argparse.ArgumentParser(description="按规则文件一键排版 Word 论文")
    ap.add_argument("input", help="要排版的 .docx 文件")
    ap.add_argument("-o", "--output", help="输出文件名，默认在原名后加 -已排版")
    ap.add_argument("-r", "--rules", default="rules.yaml", help="规则文件")
    args = ap.parse_args()

    src = Path(args.input)
    out = Path(args.output) if args.output else src.with_name(src.stem + "-已排版.docx")
    rules = load_rules(args.rules)

    stats = process(src, out, rules)

    print("=" * 52)
    print("排版完成")
    print("=" * 52)

    nzr = stats.get("normalize")
    if nzr:
        print("已重建文档结构（原文档没用 Word 样式）：")
        if nzr.get("md_heading"):
            print(f"  · {nzr['md_heading']} 个 Markdown 标题（# 开头）转成了 Word 标题")
        if nzr.get("inferred_headings"):
            print(f"  · 按字号推断出 {nzr['inferred_headings']} 个标题层级")
        if nzr.get("md_bold"):
            print(f"  · 清理 {nzr['md_bold']} 处 ** 加粗符号")
        if nzr.get("md_removed"):
            print(f"  · 删除 {nzr['md_removed']} 段 Markdown 残留（分隔线 / 代码块 / 表格分隔行）")
        if nzr.get("empty_removed"):
            print(f"  · 删除 {nzr['empty_removed']} 个空段落")
        if nzr.get("stripped_runs"):
            print(f"  · 清除 {nzr['stripped_runs']} 处手动格式，改由样式统一控制")
        if nzr.get("doc_title"):
            print("  · 开头的文档大标题改用 Title 样式")
        if nzr.get("aligned_headings"):
            print(f"  · 补齐 {nzr['aligned_headings']} 个漏掉样式的章节标题（一、二、三…）")
        if nzr.get("title"):
            print("  · 按字号识别出文档大标题，单独用 Title 样式")
        if nzr.get("ai_note"):
            print(f"  · 删除 {nzr['ai_note']} 处 AI 生成工具的署名（交稿前必须去掉）")
        if stats.get("abstract_labels"):
            print(f"  · {stats['abstract_labels']} 处「摘要/关键词」标签改为黑体顶格")
        if stats.get("code_tables"):
            print(f"  · {stats['code_tables']} 个代码块统一为等宽字体、无缩进")
        if nzr.get("tables_fixed"):
            print(f"  · {nzr['tables_fixed']} 个表格设为表头跨页重复")
        print()

    print(f"已应用样式：{'、'.join(stats['applied'])}")
    if stats["missing"]:
        print(f"文档里没有这些样式，已跳过：{'、'.join(stats['missing'])}")
    for n in stats["notes"]:
        print(f"  · {n}")
    print()
    print(f"一级标题段落：{stats['h1']}")
    print(f"二级标题段落：{stats['h2']}")
    print(f"三级标题段落：{stats['h3']}")
    print(f"正文段落：    {stats['body']}")
    if stats.get("weird"):
        print(f"未定义样式：  {stats['weird']}（这些段落引用了文档里不存在的样式，格式可能不生效）")
    print()

    if stats["h1"] == 0:
        print("[注意] 没有检测到任何「标题 1」段落。")
        print("  说明你的论文是用加粗放大手动做的标题，不是 Word 的标题样式。")
        print("  这种情况下本工具只能改正文，标题层级改不了。")
        print("  解决：在 Word 里选中标题 → 开始 → 样式 → 标题 1（或按 Ctrl+Alt+1）。")
        print()

    if stats.get("direct_ratio", 0) > 0.6:
        pct = stats["direct_ratio"] * 100
        print(f"[注意] 这份文档 {pct:.0f}% 的文字带独立字号，说明格式是逐段手动设的。")
        print("  工具改的是样式表，压不住手动格式——对这类文档只能改页面设置（页边距、页码），")
        print("  正文的字体字号行距不会变。")
        print("  要让工具生效：Word 里全选 → 开始 → 清除所有格式，")
        print("  再给标题套用「标题 1/2/3」，然后重跑一次。")
        print()

    print(f"已保存：{out}")


if __name__ == "__main__":
    main()

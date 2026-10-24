# -*- coding: utf-8 -*-
"""
文档规范化

AI 直接生成、从网页复制、业务系统导出的 Word 文档，通常有这些毛病：
    · 样式表是空的，或者段落引用了不存在的样式
    · 格式全靠「直接格式」（每个 run 带自己的字号字体）
    · Markdown 痕迹没洗干净：# 标题、**加粗**、--- 分隔线、| --- | 表格分隔行
    · 一堆连续空段落
    · 表格跨页后表头不重复

这些文档改样式表是没用的——直接格式优先级高于样式，压不住。
所以要先「重建结构」：把 Markdown 痕迹转成真标题、把字号层级推断成标题级别、
清掉直接格式、删掉垃圾段落，之后样式才落得下去。

处理顺序不能乱：
    1. 补 Heading 样式定义（否则第 2 步无处可挂）
    2. Markdown 清理（# 标题最可靠，优先用它定级）
    3. 字号推断（剩下的按字号 + 加粗猜层级）
    4. 清直接格式（必须在 2、3 之后，否则字号信息就没了）
    5. 删空段落
    6. 表格处理
"""

import re
from collections import Counter

from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

MD_H = re.compile(r"^\s{0,3}(#{1,6})\s+(.+?)\s*#*\s*$")
MD_HR = re.compile(r"^\s*([-*_])(?:\s*\1){2,}\s*$")
MD_TSEP = re.compile(r"^\s*\|?[\s:|-]*-[\s:|-]*\|?\s*$")
MD_FENCE = re.compile(r"^\s*```")
MD_QUOTE = re.compile(r"^\s*(?:>\s?)+")
MD_BOLD = re.compile(r"\*\*(.+?)\*\*|__(.+?)__")
MD_LIST = re.compile(r"^(\s*)[-*+]\s+")
# AI 生成工具的署名尾巴，交稿前必须去掉
AI_NOTE = re.compile(r"^[|｜\s]*[（(]?\s*注\s*[:：].{0,40}(AI|人工智能|模型).{0,20}[)）]?[|｜\s]*$")

# 「一、二、三、」这种章节编号（注意不带括号，（一）是列表项）
CN_NUM_HEAD = re.compile(r"^([一二三四五六七八九十百]+)[、.．]\s*\S")

# 「1」「1.1」「1.1.1」——点号个数就是层级，比字号可靠。
# 注意「1. 某条」这种带点又没后续数字的，多半是列表项，不算标题。
NUM_HEAD = re.compile(r"^(\d+(?:\.\d+)+)[\s、.．\u3000]|^(\d+)[\s\u3000]")


def level_from_number(text):
    """返回 1/2/3/4，或 None（不是编号标题）"""
    m = NUM_HEAD.match(text)
    if not m:
        return None
    if m.group(1):                       # 1.1 / 1.1.1
        return min(m.group(1).count(".") + 1, 4)
    return 1                             # 「1 」这种章号
# 参考文献条目 [1] / ［1］
REF_ITEM = re.compile(r"^\s*[\[［]\s*\d+\s*[\]］]")
# 「参考文献」标题
REF_HEAD = re.compile(r"^\s*(参考\s*文献|references?)\s*$", re.I)

# 清直接格式时要去掉的属性
DROP_PPR = ("w:spacing", "w:ind", "w:jc")
DROP_RPR = ("w:sz", "w:szCs", "w:rFonts", "w:b", "w:bCs", "w:kern", "w:highlight")


# --------------------------------------------------------------- 工具
def _drop_para(p):
    el = p._p
    parent = el.getparent()
    if parent is not None:
        parent.remove(el)


def _set_text(p, text):
    """把段落压成一个 run，避免一段里挤着好几个格式碎片"""
    runs = p.runs
    for r in runs[1:]:
        r._r.getparent().remove(r._r)
    if runs:
        runs[0].text = text
    elif text:
        p.add_run(text)


def _heading_like(text):
    """短、不以标点收尾、带编号前缀——大概率是标题"""
    if len(text) > 30:
        return False
    if text.endswith(("。", "，", "；", "：", ".", ",", ";", ":")):
        return False
    if re.match(r"^第[一二三四五六七八九十百]+[章节篇部分]", text):
        return True
    if re.match(r"^[一二三四五六七八九十]+[、.．]", text):
        return True
    if re.match(r"^\d+(\.\d+)*[\s、.．]?\s*\S", text) and len(text) <= 30:
        return True
    return False


def ensure_heading_styles(doc):
    """文档样式表可能是空的，先补出 Title 和 Heading 1-4 备用"""
    created = []
    for i in (1, 2, 3, 4):
        name = f"Heading {i}"
        try:
            doc.styles[name]
            continue
        except KeyError:
            pass
        style = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        name_el = style.element.find(qn("w:name"))
        if name_el is not None:
            name_el.set(qn("w:val"), f"heading {i}")
        created.append(name)

    try:
        doc.styles["Title"]
    except KeyError:
        style = doc.styles.add_style("Title", WD_STYLE_TYPE.PARAGRAPH)
        name_el = style.element.find(qn("w:name"))
        if name_el is not None:
            name_el.set(qn("w:val"), "Title")
        created.append("Title")

    try:
        doc.styles["Bibliography"]
    except KeyError:
        style = doc.styles.add_style("Bibliography", WD_STYLE_TYPE.PARAGRAPH)
        name_el = style.element.find(qn("w:name"))
        if name_el is not None:
            name_el.set(qn("w:val"), "Bibliography")
        created.append("Bibliography")

    return created


def base_font_size(doc):
    """全文最常用的字号，当作正文基准"""
    sizes = [r.font.size.pt for r in _all_runs(doc) if r.font.size is not None]
    if not sizes:
        return 12.0
    return Counter(sizes).most_common(1)[0][0]


def _all_runs(doc):
    from docx.text.run import Run
    for r_el in doc.element.body.iter(qn("w:r")):
        yield Run(r_el, doc)


# --------------------------------------------------------------- 1. Markdown
def clean_markdown(doc):
    stats = {"md_heading": 0, "md_bold": 0, "md_removed": 0}
    for p in list(doc.paragraphs):
        raw = p.text
        if not raw.strip():
            continue
        s = raw.strip()

        # AI 工具加在文末的署名，交稿前必须删掉
        if AI_NOTE.match(s):
            _drop_para(p)
            stats["md_removed"] += 1
            stats["ai_note"] = stats.get("ai_note", 0) + 1
            continue

        # 代码围栏
        if MD_FENCE.match(s):
            _drop_para(p)
            stats["md_removed"] += 1
            continue

        # 表格分隔行 |---|---|
        if "|" in s and MD_TSEP.match(s):
            _drop_para(p)
            stats["md_removed"] += 1
            continue

        # 分隔线 --- / *** / ___
        if MD_HR.match(s):
            _drop_para(p)
            stats["md_removed"] += 1
            continue

        # # 标题：最可靠的层级来源
        m = MD_H.match(s)
        if m and len(m.group(2)) <= 60:
            level = min(len(m.group(1)), 3)
            _set_text(p, m.group(2))
            try:
                p.style = doc.styles[f"Heading {level}"]
                stats["md_heading"] += 1
            except KeyError:
                pass
            continue

        # > 引用符号
        if MD_QUOTE.match(s):
            _set_text(p, MD_QUOTE.sub("", s))

        # - 列表符号：Word 里没有对应的项目符号时，留着反而更乱，去掉
        if MD_LIST.match(p.text):
            _set_text(p, MD_LIST.sub("", p.text))
            stats["md_list"] = stats.get("md_list", 0) + 1

        # **加粗**
        if MD_BOLD.search(p.text):
            for r in p.runs:
                if not r.text:
                    continue
                new = MD_BOLD.sub(lambda mm: mm.group(1) or mm.group(2), r.text)
                if new != r.text:
                    r.text = new
                    r.bold = True
                    stats["md_bold"] += 1

    return stats


# --------------------------------------------------------------- 2. 纯文本表格
def _is_table_row(text):
    s = text.strip()
    return s.startswith("|") and s.endswith("|") and s.count("|") >= 2


def convert_md_tables(doc):
    """AI 生成的表格粘进 Word 后，常常是「| a | b |」这样的纯文本行。
    这里把连续几行合成真表格。"""
    paras = list(doc.paragraphs)
    i, made = 0, 0

    while i < len(paras):
        if not _is_table_row(paras[i].text):
            i += 1
            continue

        j = i
        rows = []
        while j < len(paras) and _is_table_row(paras[j].text):
            body = paras[j].text.strip().strip("|")
            rows.append([c.strip() for c in body.split("|")])
            j += 1

        if len(rows) < 2:
            i = j
            continue

        cols = max(len(r) for r in rows)
        try:
            table = doc.add_table(rows=len(rows), cols=cols)
        except Exception:
            i = j
            continue

        for ri, row in enumerate(rows):
            for ci in range(cols):
                table.cell(ri, ci).text = row[ci] if ci < len(row) else ""

        try:
            table.style = doc.styles["Table Grid"]
        except KeyError:
            pass

        paras[i]._p.addprevious(table._tbl)
        for k in range(i, j):
            _drop_para(paras[k])

        made += 1
        i = j

    return made


# --------------------------------------------------------------- 3. 结构补漏
def detect_document_title(doc):
    """第一段居中、长度适中、不带句号，多半就是文档大标题，但被设成了正文。"""
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
        try:
            if p.style and p.style.name.startswith(("Heading", "Title")):
                return 0
        except Exception:
            pass
        if not (10 <= len(text) <= 45):
            return 0
        if text.endswith(("。", "，", ".", ",")):
            return 0
        if p.alignment is None or "CENTER" not in str(p.alignment):
            return 0
        try:
            p.style = doc.styles["Title"]
            return 1
        except KeyError:
            return 0
    return 0


def align_numbered_headings(doc):
    """同一篇文档里，「一、二、三、」这类章节标题如果有的挂了标题样式、
    有的还是正文，按已挂的层级把漏的补齐。

    这类情况来自 PDF 转 Word、或者从几个来源拼接的文档——
    标题样式断断续续，肉眼能看出是标题，程序只看到正文。
    """
    levels = Counter()
    missed = []
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text or len(text) > 30:
            continue
        if not CN_NUM_HEAD.match(text):
            continue
        try:
            name = p.style.name
        except Exception:
            name = ""
        if name.startswith("Heading "):
            try:
                levels[int(name.split()[-1])] += 1
            except ValueError:
                pass
        else:
            missed.append(p)

    if not levels or not missed:
        return 0

    level = levels.most_common(1)[0][0]
    n = 0
    for p in missed:
        try:
            p.style = doc.styles[f"Heading {level}"]
            n += 1
        except KeyError:
            continue
    return n


def apply_reference_block(doc):
    """找到「参考文献」标题，把它归到一级标题，后面的条目归到 Bibliography 样式。

    参考文献条目以 [1] 开头，字号又和正文/标题接近，只靠字号会被误判成标题。
    按区块处理最稳。
    """
    paras = list(doc.paragraphs)
    start = None
    for i, p in enumerate(paras):
        if REF_HEAD.match(p.text.strip()):
            start = i
            break
    if start is None:
        return 0

    try:
        paras[start].style = doc.styles["Heading 1"]
    except KeyError:
        pass

    n = 0
    for p in paras[start + 1:]:
        t = p.text.strip()
        if not t:
            continue
        if level_from_number(t) == 1:
            break                    # 撞到下一个一级标题，参考文献区结束
        try:
            p.style = doc.styles["Bibliography"]
            n += 1
        except KeyError:
            pass
    return n


def count_unstyled_numbered(doc):
    """有多少「一、」开头的段落没挂标题样式——用来判断结构是不是残缺。"""
    n = 0
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text or len(text) > 30:
            continue
        if not CN_NUM_HEAD.match(text):
            continue
        try:
            if p.style and p.style.name.startswith("Heading"):
                continue
        except Exception:
            pass
        n += 1
    return n


# --------------------------------------------------------------- 4. 标题推断
def infer_headings(doc, base_size):
    """给段落定标题层级。

    两种文档，两套规则，不能混：

    A. 有阿拉伯编号（1 / 1.1 / 1.1.1）——点号个数就是层级，最可靠。
    B. 没有编号（「一、」「方法 1」这类）——只能按字号分档，字号越大层级越浅。

    以前两套规则一起跑，互相干扰：中文编号的段落先被定成某一级，
    字号推断再从那一级往下接，结果大字号标题反被压到更深的层级。
    """
    cands = []      # (段落, 字号, 编号层级 or None)

    for p in list(doc.paragraphs):
        text = p.text.strip()
        if not text or len(text) > 60:
            continue
        if REF_ITEM.match(text) or REF_HEAD.match(text):
            continue

        already = False
        try:
            if p.style is not None and p.style.name.startswith(("Heading", "Title")):
                already = True
        except Exception:
            pass
        if already:
            continue

        sizes = [r.font.size.pt for r in p.runs if r.font.size is not None]
        size = max(sizes) if sizes else None
        bold = any(r.bold for r in p.runs)
        level = level_from_number(text)

        if level or (size is not None and size >= base_size + 1) or (bold and _heading_like(text)):
            cands.append((p, size if size is not None else base_size, level))

    if not cands:
        return {"title": 0, "headings": 0}

    numbered = [c for c in cands if c[2]]

    # ---------- A. 编号体系 ----------
    if len(numbered) >= 2:
        for p, _, level in numbered:
            try:
                p.style = doc.styles[f"Heading {level}"]
            except KeyError:
                pass

        rest = [c for c in cands if not c[2]]
        titles = 0
        if rest:
            rest.sort(key=lambda x: -x[1])
            biggest = round(rest[0][1], 1)
            group = [c for c in rest if round(c[1], 1) == biggest]
            if len(group) == 1 and biggest >= base_size * 1.4:
                try:
                    group[0][0].style = doc.styles["Title"]
                    rest.remove(group[0])
                    titles = 1
                except KeyError:
                    pass

        used = {c[2] for c in numbered}
        last = max(used) if used else 0
        n = 0
        known = {}
        for p, size in [(c[0], c[1]) for c in rest]:
            key = round(size, 1)
            if key in known:
                level = known[key]
            else:
                last = min(last + 1, 4)
                level = last
                known[key] = level
            try:
                p.style = doc.styles[f"Heading {level}"]
                n += 1
            except KeyError:
                continue
        return {"title": titles, "headings": n + len(numbered)}

    # ---------- B. 纯字号分档 ----------
    titles = 0
    biggest = max(round(c[1], 1) for c in cands)
    top = [c for c in cands if round(c[1], 1) == biggest]
    if len(top) == 1 and biggest >= base_size * 1.4:
        try:
            top[0][0].style = doc.styles["Title"]
            cands.remove(top[0])
            titles = 1
        except KeyError:
            pass

    if not cands:
        return {"title": titles, "headings": 0}

    # 已经存在的标题层级：新推断的接在它们后面，别抢前面的位置
    existing = set()
    for q in doc.paragraphs:
        try:
            nm = q.style.name
            if nm.startswith("Heading "):
                existing.add(int(nm.split()[-1]))
        except Exception:
            pass
    start = max(existing) if existing else 0

    tiers = sorted({round(c[1], 1) for c in cands}, reverse=True)
    n = 0
    for p, size, _ in cands:
        level = min(start + tiers.index(round(size, 1)) + 1, 4)
        try:
            p.style = doc.styles[f"Heading {level}"]
            n += 1
        except KeyError:
            continue
    return {"title": titles, "headings": n}


# --------------------------------------------------------------- 3. 清直接格式
def strip_direct_format(doc):
    n = 0
    for p_el in doc.element.body.iter(qn("w:p")):
        ppr = p_el.find(qn("w:pPr"))
        if ppr is not None:
            for tag in DROP_PPR:
                for el in ppr.findall(qn(tag)):
                    ppr.remove(el)
        for r_el in p_el.iter(qn("w:r")):
            rpr = r_el.find(qn("w:rPr"))
            if rpr is None:
                continue
            for tag in DROP_RPR:
                for el in rpr.findall(qn(tag)):
                    rpr.remove(el)
            n += 1
    return n


# --------------------------------------------------------------- 4. 空段落
def remove_empty_paragraphs(doc):
    n = 0
    for p in list(doc.paragraphs):
        if p.text.strip():
            continue
        el = p._p
        if el.findall(".//" + qn("w:drawing")):
            continue
        if el.findall(".//" + qn("w:pict")):
            continue
        if el.findall(".//" + qn("w:br")):
            continue
        if el.find(".//" + qn("w:sectPr")) is not None:
            continue
        _drop_para(p)
        n += 1
    return n


# --------------------------------------------------------------- 5. 表格
def fix_tables(doc):
    n = 0
    for t in doc.tables:
        if not len(t.rows):
            continue
        tr = t.rows[0]._tr
        trpr = tr.find(qn("w:trPr"))
        if trpr is None:
            trpr = OxmlElement("w:trPr")
            tr.insert(0, trpr)
        if trpr.find(qn("w:tblHeader")) is None:
            trpr.append(OxmlElement("w:tblHeader"))
        try:
            t.autofit = True
        except Exception:
            pass
        n += 1
    return n


# --------------------------------------------------------------- 总入口
def normalize(doc):
    """重建文档结构。返回一份处理报告。"""
    report = {
        "heading_styles_created": ensure_heading_styles(doc),
    }

    report.update(clean_markdown(doc))
    report["md_tables"] = convert_md_tables(doc)
    report["doc_title"] = detect_document_title(doc)
    report["references"] = apply_reference_block(doc)

    base = base_font_size(doc)
    report["base_size"] = base

    hinfo = infer_headings(doc, base)
    report["title"] = hinfo["title"]
    report["inferred_headings"] = hinfo["headings"]

    # 兜底：字号和正文一样、但带「一、二、三」编号的段落，按同类已定级的级别补齐
    report["aligned_headings"] = align_numbered_headings(doc)

    report["stripped_runs"] = strip_direct_format(doc)
    report["empty_removed"] = remove_empty_paragraphs(doc)
    report["tables_fixed"] = fix_tables(doc)
    return report

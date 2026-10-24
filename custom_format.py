# -*- coding: utf-8 -*-
"""
自定义格式解析

用户不用学语法，按学校文件上的说法写就行。一行一条，
冒号（或空格）前面是「要改什么」，后面是「改成什么样」。

   正文：仿宋 小四，1.5倍行距，首行缩进2字符
   一级标题：黑体 三号，居中
   二级标题：黑体 四号
   页边距：上2.5 下2.5 左3.0 右2.5

解析时按关键词识别，不依赖顺序，也不要求写全。
"""

import copy
import re

# 设置项：用户可能写的说法 → 内部键名
TARGETS = {
    "正文": "body", "body": "body", "段落": "body",
    "标题1": "h1", "标题一": "h1", "一级标题": "h1", "h1": "h1", "章标题": "h1",
    "标题2": "h2", "标题二": "h2", "二级标题": "h2", "h2": "h2", "节标题": "h2",
    "标题3": "h3", "标题三": "h3", "三级标题": "h3", "h3": "h3", "小节标题": "h3",
    "题注": "caption", "图题": "caption", "表题": "caption", "图表标题": "caption",
    "大标题": "title", "文档标题": "title", "封面标题": "title",
    "页边距": "page", "边距": "page", "页面": "page", "页边距设置": "page",
}

FONTS = [
    "仿宋_GB2312", "楷体_GB2312", "方正小标宋简体",
    "宋体", "仿宋", "黑体", "楷体", "微软雅黑", "等线", "幼圆",
    "Times New Roman", "Arial", "Calibri", "Cambria", "Consolas",
]

SIZES = {
    "初号": 42, "小初": 36, "一号": 26, "小一": 24, "二号": 22, "小二": 18,
    "三号": 16, "小三": 15, "四号": 14, "小四": 12, "五号": 10.5, "小五": 9,
    "六号": 7.5, "小六": 6.5,
}

ALIGNS = {
    "居中": "center", "左对齐": "left", "右对齐": "right",
    "两端对齐": "justify", "分散对齐": "justify",
}

# 句子里只起连接作用、不改变含义的词，解析前统一抹掉。
# 用户写「正文全部改为仿宋」时，「全部」和「改为」都该被忽略，只留「仿宋」。
FILLER = re.compile(
    r"设置为|调整成|采用|使用|改成|改为|设为|设成|调成|换成|变成|"
    r"全部|所有|一律|统一|整体|整篇|一起|全都|每个|"
    r"麻烦|帮我|请|把|将|应该|需要|"
    r"字体|字号|文字|大小|段落格式|"
    r"的|地|得|了|一下|用"
)

# 边距方向词
MARGIN_KEYS = {"上": "margin_top", "下": "margin_bottom",
               "左": "margin_left", "右": "margin_right"}


class FormatError(Exception):
    pass


def _blank(text, start, end):
    """把匹配到的片段抹成空格，避免同一段被重复识别"""
    return text[:start] + " " * (end - start) + text[end:]


def _take_font(body):
    for name in sorted(FONTS, key=len, reverse=True):
        idx = body.find(name)
        if idx >= 0:
            return name, _blank(body, idx, idx + len(name))
    return None, body


_SIZE_RE = re.compile(
    "(?:" + "|".join(re.escape(k) for k in sorted(SIZES, key=len, reverse=True)) + ")"
    r"(?!点|些|下|儿|点)"          # 「小一点」不是「小一」
)

# 模糊说法：没法换算成数值，得让用户给具体值
VAGUE = re.compile(r"[大小多少长短高矮][一]?[点些]|稍微|略微|差不多|大概|左右|稍微")


def _take_size(body):
    m = _SIZE_RE.search(body)
    if m:
        end = m.end()
        if body[end:end + 1] == "号":          # 「小四号」
            end += 1
        return SIZES[m.group(0)], _blank(body, m.start(), end)
    m = re.search(r"(\d+(?:\.\d+)?)\s*(?:pt|磅)", body)
    if m:
        return float(m.group(1)), _blank(body, m.start(), m.end())
    m = re.search(r"(?<![\d.])(\d{1,2}(?:\.5)?)(?![\d.]|\s*倍)", body)
    if m and float(m.group(1)) >= 5:      # 裸数字：5 以上当字号
        return float(m.group(1)), _blank(body, m.start(), m.end())
    return None, body


def _take_line_spacing(body):
    if "单倍" in body:
        i = body.find("单倍")
        end = i + 2
        if body[end:end + 2] == "行距":
            end += 2
        return 1.0, _blank(body, i, end)
    m = re.search(r"(?:行距)?\s*(\d+(?:\.\d+)?)\s*倍(?:行距)?", body)
    if m:
        return float(m.group(1)), _blank(body, m.start(), m.end())
    m = re.search(r"固定值\s*(\d+(?:\.\d+)?)\s*(?:磅|pt)?", body)
    if m:
        return float(m.group(1)), _blank(body, m.start(), m.end())
    return None, body


def _take_indent(body):
    m = re.search(r"(?:首行)?缩进\s*(\d+(?:\.\d+)?)\s*字?符?", body)
    if m:
        return float(m.group(1)), _blank(body, m.start(), m.end())
    if "不缩进" in body or "无缩进" in body:
        i = body.find("无缩进") if "无缩进" in body else body.find("不缩进")
        return 0.0, _blank(body, i, i + 3)
    return None, body


def _take_align(body):
    for word in sorted(ALIGNS, key=len, reverse=True):
        idx = body.find(word)
        if idx >= 0:
            return ALIGNS[word], _blank(body, idx, idx + len(word))
    return None, body


def _take_margins(body):
    got = {}
    for word, key in MARGIN_KEYS.items():
        m = re.search(word + r"边?距?\s*[:：]?\s*(\d+(?:\.\d+)?)", body)
        if m:
            got[key] = float(m.group(1))
            body = _blank(body, m.start(), m.end())
    if got:
        return got, body

    nums = re.findall(r"\d+(?:\.\d+)?", body)
    if len(nums) >= 4:
        keys = ["margin_top", "margin_bottom", "margin_left", "margin_right"]
        return {k: float(v) for k, v in zip(keys, nums[:4])}, ""
    return {}, body


def _find_target(text):
    """先精确匹配，匹配不到就找文本里包含的关键词（长的优先）。
    「正文采用宋体小四号」这种自然语序也能认出是「正文」，
    同时把关键词以外的字（比如「采用宋体小四号」）还回去继续解析。"""
    if text in TARGETS:
        return TARGETS[text], text
    best_key, best_word = None, ""
    for word, key in TARGETS.items():
        if word in text and len(word) > len(best_word):
            best_key, best_word = key, word
    return best_key, best_word


def parse_line(line):
    """解析一行，返回 (键名, 配置字典)"""
    text = line.strip()
    if not text:
        return None, {}, []

    m = re.match(r"^([^:：]+)[:：]\s*(.*)$", text)
    if m:
        target_text, body = m.group(1).strip(), m.group(2).strip()
    else:
        parts = re.split(r"[,，\s]+", text, maxsplit=1)
        target_text = parts[0]
        body = parts[1] if len(parts) > 1 else ""

    key, word = _find_target(target_text)
    if not key:
        raise FormatError(
            f"不认识的设置项「{target_text}」。"
            f"可写：正文、一级标题、二级标题、三级标题、题注、大标题、页边距"
        )
    if word and word != target_text:
        rest = target_text.replace(word, " ", 1).strip()
        if rest:
            body = (rest + " " + body).strip()
    if not body:
        raise FormatError(f"「{target_text}」后面没有写要改成什么")

    body = FILLER.sub(" ", body)

    cfg = {}
    warns = []

    if key == "page":
        margins, rest = _take_margins(body)
        if not margins:
            raise FormatError("页边距要写 4 个数字，或写成「上2.5 下2.5 左3 右2.5」")
        cfg.update(margins)
        if rest.strip():
            warns.append(f"已跳过「{rest.strip()}」")
        return key, cfg, warns

    # 字体 / 字号 / 行距 / 缩进 / 对齐，逐个抠，顺序无所谓
    val, body = _take_font(body)
    if val:
        cfg["font_cn"] = val

    val, body = _take_size(body)
    if val:
        cfg["size"] = val

    val, body = _take_line_spacing(body)
    if val:
        cfg["line_spacing"] = val

    val, body = _take_indent(body)
    if val is not None:
        cfg["first_line_chars"] = val

    val, body = _take_align(body)
    if val:
        cfg["align"] = val

    leftover = body.strip(" ，,。；;、\t:：")

    if not cfg:
        if VAGUE.search(leftover or text):
            raise FormatError(
                f"「{leftover or text}」这种说法没法换算成具体数值。"
                f"字号请写成 小四 或 12pt，行距写成 1.5倍，缩进写成 缩进2字符"
            )
        raise FormatError(
            f"没认出任何格式。你写的是「{leftover or text}」，"
            f"可以写成「正文：宋体 小四，1.5倍行距」这种样子"
        )
    if leftover:
        warns.append(f"已跳过「{leftover}」")

    return key, cfg, warns


def parse(text):
    """解析整段输入，返回 (覆盖项字典, 错误列表, 提示列表)

    错误 = 这一行没用上；提示 = 用上了，但有部分内容被忽略。
    """
    rules, errors, warns = {}, [], []
    for lineno, line in enumerate(text.splitlines(), 1):
        s = line.strip()
        if not s or s.startswith("#") or s.startswith("//"):
            continue
        try:
            key, cfg, line_warns = parse_line(s)
        except FormatError as e:
            errors.append(f"第 {lineno} 行：{e}")
            continue
        if key and cfg:
            rules.setdefault(key, {}).update(cfg)
        for w in line_warns:
            warns.append(f"第 {lineno} 行：{w}")
    return rules, errors, warns


def merge(base, override):
    """把自定义项盖到默认规则上，返回完整规则"""
    result = copy.deepcopy(base)
    for key, cfg in override.items():
        result.setdefault(key, {})
        result[key].update(cfg)
    return result


def summarize(override):
    """把覆盖项汇总成一句人话，显示在界面上"""
    if not override:
        return "默认"
    parts = []
    names = {"title": "大标题", "body": "正文", "h1": "一级标题",
             "h2": "二级标题", "h3": "三级标题", "caption": "题注", "page": "页边距"}
    for key, cfg in override.items():
        if key == "page":
            order = [("margin_top", "上"), ("margin_bottom", "下"),
                     ("margin_left", "左"), ("margin_right", "右")]
            vals = " ".join(f"{lab}{cfg[k]:g}" for k, lab in order if k in cfg)
            parts.append(f"页边距 {vals}" if vals else "页边距")
            continue
        bits = []
        if "font_cn" in cfg:
            bits.append(cfg["font_cn"])
        if "size" in cfg:
            bits.append(f"{cfg['size']:g}pt")
        if "line_spacing" in cfg:
            ls = cfg["line_spacing"]
            bits.append(f"固定值 {ls:g} 磅" if ls > 5 else f"{ls:g} 倍行距")
        if "first_line_chars" in cfg:
            bits.append(f"缩进 {cfg['first_line_chars']:g}")
        if "align" in cfg:
            bits.append({"center": "居中", "left": "左对齐",
                         "right": "右对齐", "justify": "两端对齐"}.get(cfg["align"], ""))
        parts.append(f"{names.get(key, key)} {' '.join(b for b in bits if b)}")
    return "；".join(parts)


SAMPLE = """正文：宋体 小四，单倍行距，首行缩进2字符
一级标题：黑体 小三
二级标题：黑体 四号
三级标题：黑体 小四
页边距：上2.54 下2.54 左3.17 右3.17"""

"""
文本格式规则解析器 —— 将中文格式描述文本解析为结构化格式规则。
处理类似 "题目：宋体，二号，加粗" 的学校格式要求文本。
"""

import re
from typing import Dict, Any, Optional

# ── 中文字号 → pt 映射 ──────────────────────────────────
# 按优先级排列：长关键字优先，"小X"优先于"X号"
CN_SIZE_TO_PT = [
    ("小初", 36), ("初号", 42),
    ("小一", 24), ("一号", 26),
    ("小二", 18), ("二号", 22),
    ("小三", 15), ("三号", 16),
    ("小四", 12), ("四号", 14),
    ("小五", 9),  ("五号", 10.5),
]


def _parse_font_size(text: str) -> Optional[float]:
    """从文本中提取字号（支持中文号和数字pt）"""
    # 中文号 —— 按优先级列表匹配（"小四"优先于"四号"）
    for cn, pt in CN_SIZE_TO_PT:
        if cn in text:
            return pt
    # 数字 pt
    m = re.search(r"(\d+(?:\.\d+)?)\s*pt", text, re.IGNORECASE)
    if m:
        return float(m.group(1))
    # 数字"号"（如 5号）
    m = re.search(r"(\d+)\s*号", text)
    if m:
        num = int(m.group(1))
        # 标准字号映射：三号=16, 四号=14, 小四=12, 五号=10.5, 小五=9
        pt_map = {1: 26, 2: 22, 3: 16, 4: 14, 5: 10.5}
        return pt_map.get(num)
    return None


def _parse_font_name(text: str) -> Optional[str]:
    """从文本中提取字体名称"""
    fonts = ["宋体", "黑体", "楷体", "仿宋", "微软雅黑",
             "Times New Roman", "Arial", "Calibri"]
    for f in fonts:
        if f in text:
            return f
    return None


def _parse_line_spacing(text: str) -> Optional[float]:
    """从文本中提取行距倍数"""
    m = re.search(r"(\d+(?:\.\d+)?)\s*倍\s*行距", text)
    if m:
        return float(m.group(1))
    # 固定值行距（如 20磅）
    m = re.search(r"(\d+)\s*磅", text)
    if m:
        return float(m.group(1)) / 12  # 近似换算
    return None


def _parse_alignment(text: str) -> Optional[str]:
    """从文本中提取对齐方式"""
    ali_map = {
        "居中": "center", "置中": "center",
        "左对齐": "left", "居左": "left",
        "右对齐": "right", "居右": "right",
        "两端对齐": "justify",
        "左侧": "left", "右侧": "right",
    }
    for cn, en in ali_map.items():
        if cn in text:
            return en
    return None


def _parse_bold(text: str) -> Optional[bool]:
    if "加粗" in text or "粗体" in text:
        return True
    return None


def _parse_first_line_indent(text: str) -> Optional[str]:
    """首行缩进"""
    m = re.search(r"首行缩进\s*(\d+)\s*字符", text)
    if m:
        chars = int(m.group(1))
        return f"{chars * 0.74:.2f}cm"  # 1字符≈0.74cm（小四）
    m = re.search(r"首行缩进\s*(\d+(?:\.\d+)?)\s*(cm|厘米)", text)
    if m:
        return f"{m.group(1)}cm"
    return None


# ── 元素名识别 ─────────────────────────────────────────
# 优先级从高到低匹配
ELEMENT_PATTERNS = [
    # (正则, 目标键, 描述)
    (r"题目|论文题目|文档标题|大标题", "Title", "文档标题"),
    (r"一级标题|1级标题|标题1", "Heading1", "一级标题"),
    (r"二级标题|2级标题|标题2", "Heading2", "二级标题"),
    (r"三级标题|3级标题|标题3", "Heading3", "三级标题"),
    (r"四级标题|4级标题|标题4", "Heading4", "四级标题"),
    # 单独的"标题"（不含数字）→ 默认 Heading1
    (r"(?<![级\d])\b标题\b(?![级\d])", "Heading1", "标题"),
    (r"正文|正文部分|正文内容|段落", "Normal", "正文"),
    (r"页眉", "Header", "页眉"),
    (r"页脚|页码", "Footer", "页脚"),
]

# 页眉页脚特殊处理
HEADER_FOOTER_PATTERNS = [
    (r"页眉[：:]\s*(.+?)(?:[（(]|$)", "header_text"),
    (r"页脚[：:]\s*(.+?)(?:[（(]|$)", "footer_text"),
    (r"页码[：:]\s*(.+)", "page_number"),
]


def parse_text_rules(text: str) -> Dict[str, Any]:
    """
    解析中文格式描述文本，返回结构化格式规则。

    示例输入：
    "题目：宋体，二号，加粗
    标题：宋体，三号，加粗
    正文部分：宋体，小四号字 ，1.5倍行距
    页眉：信息科学与技术学院 （宋体，小5号字）
    页码：页面底端 右侧"

    返回：与 format_extractor 相同结构的 format_rules
    """
    rules = {
        "source": "text_rules",
        "styles": {},
        "header_footer": {},
        "page_setup": {},
    }

    # 按行处理
    lines = text.strip().split("\n")

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # ── 识别元素类型 ──
        element_key = None
        element_desc = None

        for pattern, key, desc in ELEMENT_PATTERNS:
            if re.search(pattern, line):
                element_key = key
                element_desc = desc
                break

        if element_key is None:
            # 无法识别 → 尝试归为正文
            if any(kw in line for kw in ["宋体", "黑体", "楷体", "字体", "字号", "行距"]):
                element_key = "Normal"
                element_desc = "正文"
            else:
                continue

        # ── 处理页眉 ──
        if element_key == "Header":
            font_size = _parse_font_size(line)
            font_name = _parse_font_name(line) or "宋体"
            content_text = re.sub(r"页眉[：:]|（[^）]*）|\([^)]*\)", "", line).strip()
            # 尝试提取括号中的字体信息
            bracket_m = re.search(r"[（(]([^）)]*)[）)]", line)
            if bracket_m:
                bracket_text = bracket_m.group(1)
                bs = _parse_font_size(bracket_text)
                if bs:
                    font_size = bs
                bn = _parse_font_name(bracket_text)
                if bn:
                    font_name = bn
            rules["header_footer"]["header_text"] = content_text.strip()
            rules["header_footer"]["header_font"] = {
                "font_name": font_name,
                "font_size": font_size or 9,
            }
            continue

        # ── 处理页脚/页码 ──
        if element_key == "Footer":
            rules["header_footer"]["has_page_number"] = True
            rules["header_footer"]["page_number_position"] = \
                "right" if "右" in line else ("center" if "中" in line else "left")
            continue

        # ── 提取格式属性 ──
        font_name = _parse_font_name(line)
        font_size = _parse_font_size(line)
        line_spacing = _parse_line_spacing(line)
        alignment = _parse_alignment(line)
        bold = _parse_bold(line)
        first_line_indent = _parse_first_line_indent(line)

        # 构建样式规则
        style_rule = {}
        font_rule = {}
        para_rule = {}

        if font_name:
            font_rule["font_name"] = font_name
        if font_size:
            font_rule["font_size"] = font_size
        if bold is not None:
            font_rule["bold"] = bold

        if line_spacing:
            para_rule["line_spacing"] = line_spacing
        if alignment:
            para_rule["alignment"] = alignment
        if first_line_indent:
            para_rule["first_line_indent"] = first_line_indent

        # 标题默认居中
        if element_key == "Title" and "alignment" not in para_rule:
            para_rule["alignment"] = "center"

        if font_rule:
            style_rule["font"] = font_rule
        if para_rule:
            style_rule["paragraph"] = para_rule

        if style_rule:
            rules["styles"][element_key] = style_rule

    # ── 补全默认值 ──
    # 如果标题没定义但正文定义了，标题继承正文字体（放大）
    if "Normal" in rules["styles"] and "Heading1" not in rules["styles"]:
        normal_font = rules["styles"]["Normal"].get("font", {})
        rules["styles"]["Heading1"] = {
            "font": {
                "font_name": normal_font.get("font_name", "黑体"),
                "font_size": normal_font.get("font_size", 12) + 4,
                "bold": True,
            },
            "paragraph": {"space_before": "0.5cm", "space_after": "0.3cm"},
        }
    if "Normal" in rules["styles"] and "Heading2" not in rules["styles"]:
        normal_font = rules["styles"]["Normal"].get("font", {})
        rules["styles"]["Heading2"] = {
            "font": {
                "font_name": normal_font.get("font_name", "黑体"),
                "font_size": normal_font.get("font_size", 12) + 2,
                "bold": True,
            },
            "paragraph": {"space_before": "0.3cm", "space_after": "0.2cm"},
        }

    # 补充默认页边距
    if not rules["page_setup"]:
        rules["page_setup"] = {
            "margin_top": "2.54cm",
            "margin_bottom": "2.54cm",
            "margin_left": "3.18cm",
            "margin_right": "3.18cm",
        }

    return rules


def parse_text_rules_simple(text: str) -> Optional[Dict[str, Any]]:
    """简易入口：返回 None 表示解析失败，应由上层兜底"""
    try:
        result = parse_text_rules(text)
        if result.get("styles"):
            return result
    except Exception:
        pass
    return None

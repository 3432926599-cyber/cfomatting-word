"""
内容解析器 —— 将用户输入的文本解析为结构化内容。
主模式：纯文本启发式解析（适配学生一键复制的散乱文本）
辅助模式：Markdown 解析（适配有标记的文本）
"""

import re
from typing import List, Dict, Any, Optional


# ── 中文标题识别正则 ──────────────────────────────────
# 第X章 / 第X节
RE_CHAPTER = re.compile(r"^第[一二三四五六七八九十\d]+[章节].*")
# 中文数字编号：一、二、三、...
RE_CN_NUM = re.compile(r"^[一二三四五六七八九十]+[、，,]\s*")
# 带括号编号：(一) (二) 等
RE_CN_BRACKET = re.compile(r"^[（(][一二三四五六七八九十\d]+[）)]\s*")
# 纯数字编号：1. 1、1.1 1.1.1（空格也算分隔符）
RE_NUM = re.compile(r"^\d+(?:\.\d+)*[.)、\s]\s*")
# 常见关键章节词
SECTION_KEYWORDS = ["摘要", "关键词", "绪论", "引言", "前言",
                    "文献综述", "研究方法", "实验", "结果", "讨论",
                    "结论", "展望", "参考文献", "致谢", "附录",
                    "背景", "目的", "意义", "实习目的", "实习意义",
                    "实习内容", "实习总结", "收获与体会"]


def _is_heading_line(line: str) -> Optional[int]:
    """
    判断一行文本是否为标题，返回标题层级（1~3）或 None。
    """
    stripped = line.strip()
    if not stripped:
        return None

    # "第X章" → 一级标题
    if RE_CHAPTER.match(stripped):
        return 1

    # "一、XXX" 中文数字编号 → 一级标题
    if RE_CN_NUM.match(stripped) and len(stripped) < 60:
        return 1

    # "(一) XXX" → 二级标题
    if RE_CN_BRACKET.match(stripped) and len(stripped) < 60:
        return 2

    # "1." "1.1" "1.1.1" 数字编号 → 层级按小数点深度
    m_num = RE_NUM.match(stripped)
    if m_num and len(stripped) < 60:
        depth = m_num.group().count(".") + 1  # "1."=1, "1.1"=2, "1.1.1"=3
        return min(depth + 1, 4)  # +1因为"第X章"是L1，数字节是L2起

    # 关键章节词 + 短行 → 一级标题
    for kw in SECTION_KEYWORDS:
        if stripped.startswith(kw) and len(stripped) < 50:
            return 1

    # 短行（<30字）+ 无句末标点 → 可能标题，二级
    if len(stripped) < 30:
        has_end = bool(re.search(r"[。！？；，\.!\?,;]$", stripped))
        if not has_end:
            return 2

    return None


def _is_list_item(line: str) -> Optional[tuple]:
    """
    判断是否为列表项。返回 (ordered, text) 或 None。
    """
    stripped = line.strip()

    # 有序列表：1. / 1) / 1、
    m = re.match(r"^(\d+)[.)、]\s+(.+)", stripped)
    if m:
        return (True, m.group(2))

    # 无序列表：- / * / • / ·
    m = re.match(r"^[-*•·]\s+(.+)", stripped)
    if m:
        return (False, m.group(1))

    return None


def _is_strong_heading(line: str) -> bool:
    """
    仅匹配"确定无疑"的标题模式，用于段落合并时的断行判断。
    不能因为"短行无标点"就断句——那会把段落中间的短行也切断。
    """
    stripped = line.strip()
    if not stripped:
        return False
    # 第X章 / 第X节
    if RE_CHAPTER.match(stripped):
        return True
    # 一、二、三、...
    if RE_CN_NUM.match(stripped):
        return True
    # (一) (二)
    if RE_CN_BRACKET.match(stripped):
        return True
    # 1. / 1.1 / 1.1.1 数字编号
    if RE_NUM.match(stripped):
        return True
    # 关键章节词（整行只有这个词或很少字）
    for kw in SECTION_KEYWORDS:
        if stripped == kw or (stripped.startswith(kw) and len(stripped) <= len(kw) + 6):
            return True
    # Markdown 标题
    if re.match(r"^#{1,6}\s", stripped):
        return True
    return False


def _merge_paragraph_lines(lines: list) -> list:
    """
    合并散乱的断行：同一段落内单换行合并，空行作为段落分隔。
    返回合并后的行列表。
    """
    merged = []
    buf = []

    for line in lines:
        stripped = line.rstrip()

        # 空行 → 段落分隔
        if not stripped:
            if buf:
                merged.append(" ".join(buf))
                buf = []
            continue

        # 强标题模式 → 独立行
        if _is_strong_heading(stripped):
            if buf:
                merged.append(" ".join(buf))
                buf = []
            merged.append(stripped)
            continue

        # 列表行 → 独立行
        if _is_list_item(stripped) is not None:
            if buf:
                merged.append(" ".join(buf))
                buf = []
            merged.append(stripped)
            continue

        # Markdown 标题
        if re.match(r"^#{1,6}\s", stripped):
            if buf:
                merged.append(" ".join(buf))
                buf = []
            merged.append(stripped)
            continue

        # 普通行 → 加入缓冲区
        buf.append(stripped)

    # 剩余缓冲区
    if buf:
        merged.append(" ".join(buf))

    return merged


def parse_plain_text(text: str) -> Dict[str, Any]:
    """
    纯文本启发式解析 —— 适配学生一键复制的散乱文本。

    处理逻辑：
    1. 合并同一段落内的断行（空行分隔段落）
    2. 识别中文标题模式（第X章、一、、(一)、关键词等）
    3. 识别列表项
    4. 其余视为正文段落
    """
    raw_lines = text.strip().split("\n")
    lines = _merge_paragraph_lines(raw_lines)

    elements = []
    stats = {
        "total_chars": len(text),
        "heading_count": 0,
        "paragraph_count": 0,
    }

    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if not line:
            i += 1
            continue

        # ── Markdown 标题（兼容）──
        md_match = re.match(r"^(#{1,6})\s+(.+)$", line)
        if md_match:
            level = len(md_match.group(1))
            title_text = md_match.group(2).strip()
            is_first_title = (level == 1 and
                not any(e["type"] == "title" for e in elements))
            if is_first_title:
                elements.append({"type": "title", "level": 0, "text": title_text})
            else:
                elements.append({"type": "heading", "level": level, "text": title_text})
                stats["heading_count"] += 1
            i += 1
            continue

        # ── 标题检测 ──
        heading_level = _is_heading_line(line)
        if heading_level is not None:
            clean = line
            for pat in [RE_CHAPTER, RE_CN_NUM, RE_CN_BRACKET, RE_NUM]:
                m = pat.match(clean)
                if m:
                    break

            # "第X章" 不应被当作文档标题（封面已有标题时），强制为 H1
            is_chapter = RE_CHAPTER.match(clean)
            if stats["heading_count"] == 0 and not elements and not is_chapter:
                elements.append({"type": "title", "level": 0, "text": clean})
            else:
                elements.append({"type": "heading", "level": heading_level, "text": clean})
                stats["heading_count"] += 1
            i += 1
            continue

        # ── 列表检测 ──
        list_result = _is_list_item(line)
        if list_result is not None:
            ordered, first_text = list_result
            items = [{"text": first_text, "level": 0}]
            i += 1
            # 收集连续列表项
            while i < len(lines):
                lr = _is_list_item(lines[i])
                if lr is None:
                    break
                items.append({"text": lr[1], "level": 0})
                i += 1
            elements.append({"type": "list", "ordered": ordered, "items": items})
            continue

        # ── 默认：正文段落 ──
        elements.append({"type": "paragraph", "text": line.strip()})
        stats["paragraph_count"] += 1
        i += 1

    # ── 后处理：如果没有任何标题，尝试把第一段作为标题 ──
    if stats["heading_count"] == 0 and elements:
        first = elements[0]
        if first["type"] == "paragraph" and len(first["text"]) < 60:
            first["type"] = "title"
            first["level"] = 0
            stats["paragraph_count"] -= 1

    return {"elements": elements, "stats": stats}


def parse_markdown(text: str) -> Dict[str, Any]:
    """
    Markdown 解析 —— 当用户明确使用 Markdown 标记时使用。
    保留原有逻辑。
    """
    lines = text.strip().split("\n")
    elements = []
    stats = {"total_chars": len(text), "heading_count": 0, "paragraph_count": 0}

    i = 0
    while i < len(lines):
        line = lines[i]
        if not line.strip():
            i += 1
            continue

        # 标题
        m = re.match(r"^(#{1,6})\s+(.+)$", line)
        if m:
            level = len(m.group(1))
            txt = m.group(2).strip()
            if level == 1 and not any(e["type"] == "title" for e in elements):
                elements.append({"type": "title", "level": 0, "text": txt})
            else:
                elements.append({"type": "heading", "level": level, "text": txt})
                stats["heading_count"] += 1
            i += 1
            continue

        # 无序列表
        ul_m = re.match(r"^(\s*)[-*]\s+(.+)$", line)
        if ul_m:
            items = []
            base = len(ul_m.group(1))
            while i < len(lines):
                m2 = re.match(r"^(\s*)[-*]\s+(.+)$", lines[i])
                if not m2:
                    break
                lvl = max(0, (len(m2.group(1)) - base) // 2)
                items.append({"text": m2.group(2).strip(), "level": lvl})
                i += 1
            elements.append({"type": "list", "ordered": False, "items": items})
            continue

        # 有序列表
        ol_m = re.match(r"^(\s*)\d+[.)]\s+(.+)$", line)
        if ol_m:
            items = []
            while i < len(lines):
                m2 = re.match(r"^(\s*)\d+[.)]\s+(.+)$", lines[i])
                if not m2:
                    break
                items.append({"text": m2.group(2).strip(), "level": 0})
                i += 1
            elements.append({"type": "list", "ordered": True, "items": items})
            continue

        # 段落
        elements.append({"type": "paragraph", "text": line.strip()})
        stats["paragraph_count"] += 1
        i += 1

    return {"elements": elements, "stats": stats}


def parse_content(text: str, content_type: str = "auto") -> Dict[str, Any]:
    """
    统一入口 —— 自动选择合适的解析模式。

    content_type:
      - "auto": 自动检测（有 Markdown 标记则用 Markdown，否则纯文本）
      - "markdown": 强制 Markdown
      - "plain": 强制纯文本
    """
    if content_type == "markdown":
        return parse_markdown(text)

    if content_type == "plain":
        return parse_plain_text(text)

    # auto: 检测是否有明显的 Markdown 标记
    has_md = bool(re.search(r"^#{1,6}\s", text, re.MULTILINE))
    # 检测是否有占比较高的 md 标记
    lines = text.split("\n")
    md_lines = sum(1 for l in lines if re.match(r"^#{1,6}\s", l))
    if md_lines >= 2 or (has_md and md_lines / max(len(lines), 1) > 0.05):
        return parse_markdown(text)

    return parse_plain_text(text)

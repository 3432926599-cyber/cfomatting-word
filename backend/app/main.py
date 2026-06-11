"""
CFomatting-word 后端入口
轻量级文档格式化工具 —— 上传范本 + 粘贴内容 → 生成 Word
"""

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import os
import uuid
import json
import shutil
from datetime import datetime


def _convert_doc_to_docx(doc_path: str) -> str:
    """
    将旧版 .doc 转换为 .docx。
    Windows → Word COM；Linux → LibreOffice headless。
    """
    abs_path = os.path.abspath(doc_path)
    docx_path = doc_path + "x"

    # 方法1: Word COM (Windows)
    try:
        import pythoncom
        pythoncom.CoInitialize()
        try:
            import win32com.client
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            try:
                doc = word.Documents.Open(abs_path)
                doc.SaveAs2(os.path.abspath(docx_path), FileFormat=16)
                doc.Close()
                return docx_path
            finally:
                word.Quit()
        finally:
            pythoncom.CoUninitialize()
    except (ImportError, Exception):
        pass

    # 方法2: LibreOffice headless (Linux / Render)
    import subprocess
    out_dir = os.path.dirname(abs_path)
    result = subprocess.run(
        ["libreoffice", "--headless", "--convert-to", "docx",
         "--outdir", out_dir, abs_path],
        capture_output=True, text=True, timeout=60,
    )
    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice conversion failed: {result.stderr}")
    if os.path.exists(docx_path):
        return docx_path
    raise RuntimeError("LibreOffice did not produce output file")

from app.format_extractor import extract_format
from app.content_parser import parse_content
from app.text_rule_parser import parse_text_rules
from app.doc_generator import generate_docx

# ── 路径配置 ──────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")
STATIC_DIR = os.path.join(BASE_DIR, "static")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(STATIC_DIR, exist_ok=True)

# ── FastAPI 应用 ─────────────────────────────────────
app = FastAPI(
    title="CFomatting-word",
    description="文档格式自动化工具 — 上传范本 + 粘贴内容 → 生成规范 Word 文档",
    version="0.1.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── 会话存储（简易内存存储，MVP 够用）───────────────
sessions: dict = {}  # {session_id: {format_rules, created_at, ...}}


def _get_or_create_session(session_id: str = None) -> str:
    """获取或创建会话"""
    if session_id and session_id in sessions:
        return session_id
    new_id = str(uuid.uuid4())[:8]
    sessions[new_id] = {
        "format_rules": None,
        "created_at": datetime.now().isoformat(),
    }
    return new_id


# ═══════════════════════════════════════════════════════
# API 端点
# ═══════════════════════════════════════════════════════


@app.post("/api/upload-sample")
async def upload_sample(file: UploadFile = File(...), session_id: str = Form("")):
    """
    上传格式范本 .docx，提取格式规则。
    返回提取的格式规则 JSON。
    """
    # 验证文件类型
    if not file.filename.endswith((".docx", ".doc")):
        raise HTTPException(400, "仅支持 .docx / .doc 文件")

    # 创建会话
    sid = _get_or_create_session(session_id)

    # 保存上传文件
    ext = os.path.splitext(file.filename)[1]
    save_path = os.path.join(UPLOAD_DIR, f"{sid}_sample{ext}")
    with open(save_path, "wb") as f:
        content = await file.read()
        f.write(content)

    # 如果是旧版 .doc，先转换为 .docx
    extract_path = save_path
    is_old_doc = ext.lower() == ".doc"
    if is_old_doc:
        try:
            extract_path = _convert_doc_to_docx(save_path)
        except Exception as e:
            raise HTTPException(500, f"旧版 .doc 转换失败（请用 Word 另存为 .docx 后重试）: {str(e)}")

    # 提取格式
    try:
        format_rules = extract_format(extract_path)

        # 智能回退：如果样本提取质量太差（没有标题样式），
        # 尝试读取文档文本，用文本规则解析器再解析一次
        styles = format_rules.get("styles", {})
        heading_keys = {"Title", "Heading1", "Heading2", "Heading3"}
        has_headings = bool(heading_keys & set(styles.keys()))

        if not has_headings:
            # 智能回退：从文档中提取纯文本，尝试用文本规则解析
            # 适用于学校给的"格式说明文档"（里面写的是格式要求文字而非格式化范本）
            try:
                # 用 python-docx 从已转换的 .docx 读取所有文本
                from docx import Document as DocxReader
                docx_doc = DocxReader(extract_path)
                doc_text = "\n".join(p.text for p in docx_doc.paragraphs if p.text.strip())

                # 尝试用文本规则解析
                from app.text_rule_parser import parse_text_rules
                text_rules = parse_text_rules(doc_text)
                if text_rules.get("styles") and len(text_rules["styles"]) >= 2:
                    format_rules = text_rules
                    sessions[sid]["format_source"] = "text_from_doc"
            except Exception:
                pass  # 回退失败，用原结果

        sessions[sid]["format_rules"] = format_rules
        sessions[sid]["sample_name"] = file.filename
        # 保存模板路径，供生成时保留封面
        sessions[sid]["template_path"] = extract_path
    except Exception as e:
        raise HTTPException(500, f"格式提取失败: {str(e)}")

    return JSONResponse(
        {
            "success": True,
            "session_id": sid,
            "sample_name": file.filename,
            "format_rules": format_rules,
        }
    )


@app.post("/api/parse-text-rules")
async def parse_text_rules_endpoint(data: dict):
    """
    解析中文格式描述文本 → 结构化格式规则。
    请求体：
    {
        "session_id": "xxx",
        "text": "题目：宋体，二号，加粗\n标题：宋体，三号，加粗\n正文：宋体，小四，1.5倍行距"
    }
    返回：{"success": true, "format_rules": {...}}
    """
    session_id = data.get("session_id", "")
    text = data.get("text", "")

    if not text.strip():
        raise HTTPException(400, "格式描述文本不能为空")

    sid = _get_or_create_session(session_id)

    try:
        format_rules = parse_text_rules(text)
        sessions[sid]["format_rules"] = format_rules
        sessions[sid]["format_source"] = "text_rules"
    except Exception as e:
        raise HTTPException(500, f"格式解析失败: {str(e)}")

    return JSONResponse({
        "success": True,
        "session_id": sid,
        "format_rules": format_rules,
        "parsed_styles": list(format_rules.get("styles", {}).keys()),
    })


@app.post("/api/generate")
async def generate_document(data: dict):
    """
    根据格式规则 + 内容，生成 Word 文档。
    请求体：
    {
        "session_id": "xxx",
        "content": "用户粘贴的文本（纯文本/Markdown均可）",
        "content_type": "auto" | "markdown" | "plain",
        "text_rules": "题目：宋体，二号，加粗\\n正文：宋体，小四，1.5倍行距"  // 可选，直接传入格式文本
    }
    返回：{"success": true, "file_id": "xxx", "download_url": "/api/download/xxx"}
    """
    session_id = data.get("session_id", "")
    content_text = data.get("content", "")
    content_type = data.get("content_type", "auto")
    text_rules = data.get("text_rules", "")

    if not content_text.strip():
        raise HTTPException(400, "内容不能为空")

    # 获取格式规则
    sid = _get_or_create_session(session_id)

    # 格式规则来源：text_rules + 样本规则合并
    format_rules = None
    sample_rules = sessions[sid].get("format_rules")  # 上传样本时提取的
    text_parsed = None

    if text_rules and text_rules.strip():
        try:
            text_parsed = parse_text_rules(text_rules)
        except Exception:
            pass

    if sample_rules and text_parsed:
        # 混合模式：样本规则打底 + 文本规则覆盖
        import copy
        format_rules = copy.deepcopy(sample_rules)
        format_rules["source"] = "hybrid"
        # 文本规则覆盖同名字段
        if text_parsed.get("styles"):
            for k, v in text_parsed["styles"].items():
                format_rules["styles"][k] = v
        if text_parsed.get("header_footer"):
            format_rules.setdefault("header_footer", {}).update(text_parsed["header_footer"])
        if text_parsed.get("page_setup"):
            format_rules.setdefault("page_setup", {}).update(text_parsed["page_setup"])
        sessions[sid]["format_rules"] = format_rules
        sessions[sid]["format_source"] = "hybrid"
    elif text_parsed:
        format_rules = text_parsed
        sessions[sid]["format_rules"] = format_rules
        sessions[sid]["format_source"] = "text_rules"
    elif sample_rules:
        format_rules = sample_rules

    if format_rules is None:
        # 使用默认学术格式（参考学校通用要求）
        format_rules = {
            "page_setup": {
                "margin_top": "2.54cm",
                "margin_bottom": "2.54cm",
                "margin_left": "3.18cm",
                "margin_right": "3.18cm",
            },
            "styles": {
                "Title": {
                    "font": {"font_name": "宋体", "font_size": 22, "bold": True},
                    "paragraph": {"alignment": "center"},
                },
                "Heading1": {
                    "font": {"font_name": "宋体", "font_size": 16, "bold": True},
                    "paragraph": {"space_before": "0.5cm", "space_after": "0.3cm"},
                },
                "Heading2": {
                    "font": {"font_name": "宋体", "font_size": 14, "bold": True},
                    "paragraph": {"space_before": "0.3cm", "space_after": "0.2cm"},
                },
                "Heading3": {
                    "font": {"font_name": "宋体", "font_size": 12, "bold": True},
                    "paragraph": {"space_before": "0.2cm", "space_after": "0.1cm"},
                },
                "Normal": {
                    "font": {"font_name": "宋体", "font_size": 12},
                    "paragraph": {"line_spacing": 1.5, "first_line_indent": "0.74cm"},
                },
            },
        }

    # 解析内容（auto: 自动识别纯文本/Markdown）
    content_data = parse_content(content_text, content_type)

    # 生成文档
    file_id = str(uuid.uuid4())[:8]
    output_path = os.path.join(OUTPUT_DIR, f"{file_id}.docx")

    # 获取模板路径（如果有上传样本，则保留封面）
    template_path = sessions[sid].get("template_path", "")
    if template_path and not os.path.exists(template_path):
        template_path = ""

    try:
        generate_docx(content_data, format_rules, output_path, template_path or None)
    except Exception as e:
        raise HTTPException(500, f"文档生成失败: {str(e)}")

    return JSONResponse(
        {
            "success": True,
            "file_id": file_id,
            "download_url": f"/api/download/{file_id}",
            "stats": content_data.get("stats", {}),
        }
    )


@app.get("/api/download/{file_id}")
async def download_document(file_id: str):
    """下载生成的 Word 文档"""
    file_path = os.path.join(OUTPUT_DIR, f"{file_id}.docx")
    if not os.path.exists(file_path):
        raise HTTPException(404, "文件不存在或已过期，请重新生成")
    return FileResponse(
        file_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename="formatted_document.docx",
    )


@app.get("/api/session/{session_id}")
async def get_session(session_id: str):
    """获取会话状态"""
    if session_id not in sessions:
        return JSONResponse({"exists": False})
    return JSONResponse(
        {
            "exists": True,
            "has_sample": sessions[session_id].get("format_rules") is not None,
            "sample_name": sessions[session_id].get("sample_name"),
        }
    )


@app.get("/api/health")
async def health():
    return {"status": "ok", "version": "0.1.0"}


# ── 静态文件（前端页面） ──────────────────────────────
app.mount("/", StaticFiles(directory=STATIC_DIR, html=True), name="static")

<p align="center">
  <img src="https://img.shields.io/badge/status-live-brightgreen?style=flat" alt="status">
  <img src="https://img.shields.io/badge/python-3.13-blue?style=flat" alt="python">
  <img src="https://img.shields.io/badge/license-MIT-green?style=flat" alt="license">
</p>

# CFomatting — 文档格式自动化

> 上传格式范本 + 粘贴教师要求 + 粘贴内容 → 一键生成规范 Word 文档

一个轻量级文档格式化工具。你只需要上传一份格式范本（学校/公司的 `.docx` 模板），粘贴你的论文或报告内容，系统自动提取格式规则并生成符合规范的 Word 文档。封面、页眉、页码、字体、行距……原样保留。

---

## ✨ 功能

| 功能 | 说明 |
|------|------|
| 📄 **范本格式提取** | 上传 `.docx` / `.doc` 模板，自动提取字体、字号、行距、对齐、页眉页脚 |
| 📝 **教师要求解析** | 粘贴中文格式描述（如"正文：宋体 12pt 1.5倍行距"），自动解析为格式规则 |
| 📋 **纯文本内容识别** | 自动识别章节标题结构：`第X章` `一、` `（一）` `1.1` Markdown `#` |
| 🎨 **智能混合模式** | 范本规则打底 + 教师要求覆盖 → 最精准的格式控制 |
| 📥 **一键下载 .docx** | 生成的 Word 文档保留封面、页眉、页码，直接可用 |

---

## 🚀 快速开始

### 1. 启动后端

```bash
# 双击运行（Windows）
start.bat

# 或手动启动
cd backend
pip install -r requirements.txt
uvicorn app.main:app --host 0.0.0.0 --port 8000
```

### 2. 打开前端

- **线上版**：[https://cfomatting-word.pages.dev](https://cfomatting-word.pages.dev)
- **本地版**：[http://localhost:8000](http://localhost:8000)

启动 `start.bat` 后会自动生成公网隧道，把链接发给别人就能用。

---

## 📁 项目结构

```
CFomatting-word/
├── backend/
│   ├── app/
│   │   ├── main.py              # FastAPI 入口，API 端点
│   │   ├── format_extractor.py  # 从范本 .docx 提取格式规则
│   │   ├── content_parser.py    # 解析用户粘贴的文本内容
│   │   ├── text_rule_parser.py  # 解析中文格式描述文本
│   │   └── doc_generator.py     # 生成 Word 文档
│   ├── requirements.txt
│   ├── uploads/                 # 上传的范本文件
│   └── outputs/                 # 生成的 Word 文档
├── frontend/
│   └── index.html               # 单页前端（微信原生风格）
├── start.bat                    # 一键启动脚本
├── cloudflared.exe              # Cloudflare 隧道工具
└── backend-url.txt              # GitHub 动态指针（自动更新公网地址）
```

---

## 🔧 技术栈

| 层 | 技术 |
|----|------|
| 后端 | Python 3.13 / FastAPI / uvicorn |
| 文档处理 | python-docx / lxml |
| 前端 | 原生 HTML/CSS/JS，微信原生风格 |
| 部署 | Cloudflare Pages（前端）/ 本地 (后端) |
| 内网穿透 | Cloudflare Tunnel / cloudflared |

---

## 📱 移动端适配

前端按照微信内置浏览器标准做了完整适配：

- 微信 X5 内核动态 `--vh` 修复
- 安全区刘海屏适配
- 44px 最小触摸目标
- 16px 最小字号防 iOS 缩放
- 4 级响应式断点（480 / 768 / 960 / 1200px）

---

## 🔗 API 端点

| 方法 | 路径 | 说明 |
|------|------|------|
| `POST` | `/api/upload-sample` | 上传格式范本，返回格式规则 JSON |
| `POST` | `/api/parse-text-rules` | 解析中文格式描述文本 |
| `POST` | `/api/generate` | 生成 Word 文档 |
| `GET` | `/api/download/{file_id}` | 下载生成的 .docx |
| `GET` | `/api/health` | 健康检查 |

---

## 👤 作者

- GitHub: [@3432926599-cyber](https://github.com/3432926599-cyber)
- 项目地址: [cfomatting-word](https://github.com/3432926599-cyber/cfomatting-word)

---

## 📄 许可

MIT License

# DocuGenius Converter Skill

> 为 Claude Code 添加双向文档转换能力

[![Claude Code](https://img.shields.io/badge/Claude_Code-Skill-purple.svg)](https://github.com/anthropics/claude-code)
[![Python](https://img.shields.io/badge/Python-3.6+-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

**DocuGenius Converter** 是一个 [Claude Code](https://github.com/anthropics/claude-code) Skill，为 Claude 添加文档格式转换能力。安装后，Claude 可以直接处理 Word、Excel、PowerPoint、PDF 等多种文档格式，将它们转换为 AI 友好的 Markdown 格式进行分析。

## 为什么需要这个 Skill？

Claude Code 默认只能直接读取 Markdown 和纯文本文件。当你上传 Word、Excel、PDF 等文档时，Claude 无法直接理解其内容。

这个 Skill 解决了这个问题：

- 让 Claude 能够"阅读" Office 文档和 PDF
- 将文档转换为 Claude 易于分析的 Markdown 格式
- 支持双向转换（Markdown 转 Word）

## 功能特性

### 文档转换能力

| 转换方向                         | 支持格式                  | 特性                     |
| -------------------------------- | ------------------------- | ------------------------ |
| **Office/PDF → Markdown** | .docx, .xlsx, .pptx, .pdf | 标题、列表、表格、格式   |
| **Markdown → Word**       | .md                       | 专业文档、样式保留       |

### 智能内容提取

- **标题识别**：自动识别 Word 标题层级（Heading 1-6）及中文标题样式
- **格式保留**：保留粗体、斜体等文本格式
- **表格转换**：智能转换表格为 Markdown 格式
- **列表支持**：支持有序列表、无序列表及多级嵌套

### 轻量高效

- 总依赖仅 10-15MB
- 按需加载，只检测当前文件类型所需的库
- 智能缓存，避免重复检查依赖

## 安装

### 1. 环境要求

- **Python 3.6+**（必需）
- **Node.js 14+**（可选，仅 Markdown → Word 需要）

### 2. 安装 Skill

将整个项目复制到你的 Claude Code skills 目录：

```bash
# macOS/Linux
git clone https://github.com/yourusername/docugenius-converter-skill.git ~/.claude/skills/docugenius-converter

# Windows
git clone https://github.com/yourusername/docugenius-converter-skill.git %USERPROFILE%\.claude\skills\docugenius-converter
```

或直接下载 ZIP 解压到 skills 目录。

### 3. 安装依赖

首次使用 Skill 时，Claude Code 会自动检测依赖。如缺少依赖，会引导你安装：

```bash
pip install -r requirements.txt

# 可选：Markdown → Word 功能需要 Node.js
cd scripts/md_to_docx
npm install
```

或运行安装脚本（推荐）：

```bash
# macOS/Linux
chmod +x install.sh && ./install.sh

# Windows
install.bat
```

## 使用方式

安装 Skill 后，在 Claude Code 中直接与 Claude 对话即可：

### 分析文档

```
帮我分析这个 report.docx 的内容
```

```
读取这个 data.xlsx 并总结数据
```

### 格式转换

```
把这个 document.pdf 转成 markdown
```

```
把 notes.md 导出为 Word 文档
```

### 批量处理

```
转换 docs 文件夹里的所有文件
```

## 工作原理

当你在 Claude Code 中请求处理文档时：

1. **检测需求**：Claude 识别到需要处理 Office/PDF/Markdown 文件
2. **调用转换脚本**：运行 `python scripts/convert_document.py`
3. **自动处理**：
   - 检测 Python 环境
   - 验证文件格式
   - 检查所需依赖（按需检查）
   - 执行转换
4. **返回结果**：返回 JSON 格式的转换结果，包含 Markdown 内容和输出路径
5. **后续处理**：Claude 分析转换后的内容，回答你的问题

### 输出位置

转换后的文件保存在原文件同级目录：

```
your-project/
├── report.docx              # 原文件
├── document.md              # 原文件
├── Markdown/                # Office/PDF → Markdown 输出
│   └── report.md
└── Word/                    # Markdown → Word 输出
    └── document.docx
```

## 支持的格式

| 格式               | 输入 | 输出 | 质量       |
| ------------------ | ---- | ---- | ---------- |
| Word (.docx)       | ✅   | ✅   | 优秀       |
| Excel (.xlsx)      | ✅   | ❌   | 优秀       |
| PowerPoint (.pptx) | ✅   | ❌   | 良好       |
| PDF (.pdf)         | ✅   | ❌   | 取决于类型 |
| Markdown (.md)     | ✅   | ✅   | 优秀       |

> **注意**：不支持旧版格式（.doc, .xls, .ppt），请先转换为新格式。

## 常见问题

### 依赖缺失怎么办？

Skill 会自动检测依赖，如果缺失会提示安装命令：

```bash
pip install python-docx    # Word
pip install openpyxl       # Excel
pip install python-pptx    # PowerPoint
pip install pdfplumber     # PDF
```

### 文件过大怎么办？

当前限制为 100MB，建议分割文件或压缩内容。

### Markdown 转 Word 失败？

需要安装 Node.js 和依赖：

```bash
cd scripts/md_to_docx
npm install
```

## 最佳实践

1. **使用新版 Office 格式**（.docx, .xlsx, .pptx）
2. **PDF 优先使用文本型**，扫描型建议先 OCR
3. **文件大小建议 < 50MB**

## 项目结构

```
docugenius-converter-skill/
├── SKILL.md                 # Claude Code Skill 定义
├── install.sh               # 安装脚本 (macOS/Linux)
├── install.bat              # 安装脚本 (Windows)
├── requirements.txt         # Python 依赖
├── scripts/
│   ├── convert_document.py  # 转换脚本核心
│   └── md_to_docx/         # Markdown → Word 模块
├── references/
│   └── supported-formats.md # 格式支持详情
└── tests/                   # 测试文件
```

## 技术细节

### 依赖管理

| 依赖        | 大小 | 用途            |
| ----------- | ---- | --------------- |
| python-docx | ~2MB | Word 处理       |
| openpyxl    | ~3MB | Excel 处理      |
| python-pptx | ~2MB | PowerPoint 处理 |
| pdfplumber  | ~5MB | PDF 处理        |

### 缓存机制

- 位置：`.claude/cache/docugenius-converter/`
- 策略：永不过期，`requirements.txt` 变更时自动失效
- 清除：`python scripts/convert_document.py --clear-cache`

## 许可证

MIT License

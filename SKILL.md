---
name: docugenius-converter
description: 双向文档转换工具，将 Word (.docx)、Excel (.xlsx)、PowerPoint (.pptx) 和 PDF (.pdf) 转换为 AI 友好的 Markdown 格式，或将 Markdown (.md) 转换为 Word (.docx) 格式。当用户请求以下操作时使用：(1) 明确请求文档转换，包括任何包含"转换"、"转为"、"转成"、"convert"、"导出"、"export"等词汇的请求（例如："转换文档"、"把这个文件转为docx"、"convert to markdown"、"导出为Word"）；(2) 需要 AI 理解文档内容（"帮我分析这个 Word 文件"、"读取这个 PDF"、"总结这个 Excel"）；(3) 上传文档文件并询问内容（"这是什么"、"帮我看看"）；(4) 任何涉及 .docx、.xlsx、.pptx、.pdf、.md 文件格式转换的请求。
---
# DocuGenius Document Converter

双向文档转换工具：

- **Office/PDF → Markdown**：将 Office 文档和 PDF 转换为 AI 友好的 Markdown 格式
- **Markdown → Word**：将 Markdown 文件转换为专业的 Word 文档

## 环境要求

- **Python 3.6+**（必需）：用于 Office/PDF 转 Markdown
- **Node.js 14+**（可选）：仅用于 Markdown 转 Word

## 核心工作流程

### 1. 检测 Python 环境

**在执行转换前，必须先检测 Python 环境**：

```bash
# Windows
python --version

# macOS/Linux
python3 --version
```

**处理结果**：

- **命令成功**：继续转换
- **命令失败**：告知用户安装 Python，提供下载链接 https://www.python.org/downloads/

### 2. 验证文件格式

支持的格式：

- **Office/PDF → Markdown**: `.docx`, `.xlsx`, `.pptx`, `.pdf`
- **Markdown → Word**: `.md`
- **不支持**: `.doc`, `.xls`, `.ppt`（旧格式需先转换）

### 3. 执行转换

```bash
python scripts/convert_document.py <file_path> [extract_images] [output_dir]
```

参数：

- `file_path`: 文档路径（必需）
- `extract_images`: `true`/`false`（默认 `true`，仅用于 Office/PDF）
- `output_dir`: 输出目录（Office/PDF 默认 `Markdown/`，Markdown 默认 `Word/`）

### 4. 处理结果

脚本输出 JSON：

```json
{
  "success": true,
  "markdown_content": "转换后的内容...",
  "output_path": "/path/to/output.md"
}
```

## 常见模式

### 模式 1: Office/PDF 转 Markdown

```bash
python scripts/convert_document.py /path/to/report.docx
```

### 模式 2: Markdown 转 Word

```bash
python scripts/convert_document.py /path/to/document.md
```

**注意**：需要 Node.js 环境。如未安装，会提示错误。

### 模式 3: 批量处理

```bash
python scripts/convert_document.py --batch /path/to/documents
```

## 依赖安装

### 方式 1: 运行安装脚本

```bash
# Windows
install.bat

# macOS/Linux
chmod +x install.sh && ./install.sh
```

### 方式 2: 手动安装

```bash
pip install -r requirements.txt
```

### 方式 3: 按需安装

```bash
pip install python-docx  # Word
pip install openpyxl     # Excel
pip install python-pptx  # PowerPoint
pip install pdfplumber   # PDF
```

## 错误处理

| 错误信息               | 原因                          | 解决方案                 |
| ---------------------- | ----------------------------- | ------------------------ |
| 缺少依赖库: xxx        | 未安装 Python 库              | 运行 `pip install xxx` |
| 文件不存在             | 路径错误                      | 检查文件路径             |
| 文件过大               | 超过 100MB 限制               | 分割文件或压缩内容       |
| 不支持的文件格式: .doc | 旧版 Office 格式              | 另存为 .docx/.xlsx/.pptx |
| 未找到 Node.js         | Markdown 转 Word 需要 Node.js | 安装 Node.js             |

## 支持的格式详情

详细的格式支持信息请参见 [supported-formats.md](references/supported-formats.md)。

快速参考：

| 格式               | 支持内容                          | 质量            |
| ------------------ | --------------------------------- | --------------- |
| Word (.docx)       | 文本、标题、列表、表格、粗体/斜体 | 优秀            |
| Excel (.xlsx)      | 表格、多工作表                    | 优秀            |
| PowerPoint (.pptx) | 幻灯片文本                        | 良好            |
| PDF (.pdf)         | 文本、表格                        | 取决于 PDF 类型 |
| Markdown (.md)     | 标题、列表、表格、代码块、格式    | 优秀            |

## 最佳实践

1. 使用新版 Office 格式（.docx/.xlsx/.pptx）
2. 文件大小建议 < 50MB
3. Word 文档使用内置 Heading 1-6 样式
4. PDF 优先使用文本型，扫描型建议先 OCR
5. 转换后检查输出质量

## 系统要求

- **Python**: 3.6+（必需）
- **Node.js**: 14+（可选，仅 Markdown 转 Word）
- **操作系统**: Windows、macOS、Linux
- **磁盘空间**: 约 50MB

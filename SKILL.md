---
name: docugenius-converter
description: 双向文档转换工具，将 Word (.docx)、Excel (.xlsx)、PowerPoint (.pptx) 和 PDF (.pdf) 转换为 AI 友好的 Markdown 格式，或将 Markdown (.md) 转换为 Word (.docx) 格式。当用户请求以下操作时使用：(1) 明确请求文档转换，包括任何包含"转换"、"转为"、"转成"、"convert"、"导出"、"export"等词汇的请求（例如："转换文档"、"把这个文件转为docx"、"convert to markdown"、"导出为Word"）；(2) 需要 AI 理解文档内容（"帮我分析这个 Word 文件"、"读取这个 PDF"、"总结这个 Excel"）；(3) 上传文档文件并询问内容（"这是什么"、"帮我看看"）；(4) 任何涉及 .docx、.xlsx、.pptx、.pdf、.md 文件格式转换的请求。
---
# DocuGenius Document Converter

双向文档转换工具：

- **Office/PDF → Markdown**：将 Office 文档和 PDF 转换为 AI 友好的 Markdown 格式
  - 输出位置：文件同级目录下的 `Markdown/` 文件夹
- **Markdown → Word**：将 Markdown 文件转换为专业的 Word 文档
  - 输出位置：文件同级目录下的 `Word/` 文件夹

## 环境要求

- **Python 3.6+**（必需）：用于 Office/PDF 转 Markdown
- **Node.js 14+**（可选）：仅用于 Markdown 转 Word

## 核心工作流程

### 1. 检测 Python 环境

使用 Bash 工具检测 Python：

```bash
# Windows
python --version

# macOS/Linux
python3 --version
```

**根据结果采取行动**：
- **成功**：继续步骤 2
- **失败**：告知用户安装 Python（https://www.python.org/downloads/），等待用户安装后再继续

### 2. 验证文件格式

检查文件扩展名：
- **支持**: `.docx`, `.xlsx`, `.pptx`, `.pdf`, `.md`
- **不支持**: `.doc`, `.xls`, `.ppt`（告知用户需先转换为新格式）

### 3. 执行转换

使用 Bash 工具运行转换脚本：

```bash
# 使用便捷脚本（自动激活全局虚拟环境）
./convert.sh <file_path> [extract_images] [output_dir]

# 或手动激活全局虚拟环境后运行
source ~/.claude/venvs/docugenius-converter/bin/activate  # Windows: %USERPROFILE%\.claude\venvs\docugenius-converter\Scripts\activate
python scripts/convert_document.py <file_path> [extract_images] [output_dir]
```

**参数说明**：
- `file_path`: 文档路径（必需）
- `extract_images`: `true`/`false`（默认 `true`，仅用于 Office/PDF）
- `output_dir`: 输出目录（可选，不指定时使用默认位置）

**依赖说明**：脚本自动检测依赖，直接执行转换即可。缺少依赖时会有错误提示。

**默认输出目录**（不指定 `output_dir` 时）：
- **Office/PDF → Markdown**: 文件同级目录下的 `Markdown/` 文件夹
  - 例如：`/path/to/document.docx` → `/path/to/Markdown/document.md`
- **Markdown → Word**: 文件同级目录下的 `Word/` 文件夹
  - 例如：`/path/to/document.md` → `/path/to/Word/document.docx``

### 4. 处理转换结果

解析 JSON 输出：
- **success: true**: 转换成功，继续步骤 5
- **success: false**: 检查 error 字段，根据错误类型处理（参见"错误处理"章节）

### 5. 向用户展示结果

转换成功后：
1. 告知用户输出文件路径
2. 如果用户请求分析内容，使用 `markdown_content` 或读取输出文件
3. 根据用户原始请求继续处理（分析、总结等）

## 常见模式

### 模式 1: 转换并分析文档

用户："分析这个 Word 文件的内容"

执行转换，解析返回的 JSON 中的 `markdown_content`，分析后向用户展示结果。

```bash
./convert.sh /path/to/report.docx
```

### 模式 2: Markdown 转 Word

用户："把这个 md 转成 docx"

执行转换，告知用户输出文件路径。

```bash
./convert.sh /path/to/document.md
```

### 模式 3: 批量转换

用户："转换这个文件夹里的所有文档"

```bash
./convert.sh --batch /path/to/documents
```

## 依赖安装

当转换脚本报错"缺少依赖库"时，引导用户安装：

### 方式 1: 运行安装脚本（推荐）

```bash
# Windows
install.bat

# macOS/Linux
chmod +x install.sh && ./install.sh
```

### 方式 2: 使用全局共享虚拟环境（推荐）

安装脚本会自动创建全局虚拟环境 `~/.claude/venvs/docugenius-converter`，所有项目共享此环境，无需重复安装。

**手动创建全局虚拟环境**：

```bash
# 创建并激活全局虚拟环境
python3 -m venv ~/.claude/venvs/docugenius-converter
source ~/.claude/venvs/docugenius-converter/bin/activate

# 安装依赖
pip install -r requirements.txt
```

### 方式 3: 按需安装特定库

```bash
# 激活全局虚拟环境后
source ~/.claude/venvs/docugenius-converter/bin/activate

pip install python-docx  # Word
pip install openpyxl     # Excel
pip install python-pptx  # PowerPoint
pip install pdfplumber   # PDF
```

## 错误处理

根据错误信息采取相应行动：

| 错误信息               | 原因               | 解决方案                                    |
| ---------------------- | ------------------ | ------------------------------------------- |
| 缺少依赖库: xxx        | 未安装 Python 库   | 使用 Bash 运行 `pip install xxx`            |
| 文件不存在             | 路径错误           | 验证文件路径，使用绝对路径                  |
| 文件过大               | 超过 100MB 限制    | 告知用户需分割文件或压缩内容                |
| 不支持的文件格式: .doc | 旧版 Office 格式   | 告知用户需先转换为 .docx/.xlsx/.pptx        |
| 未找到 Node.js         | Markdown 转 Word   | 告知用户需安装 Node.js (https://nodejs.org) |

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

## 缓存机制

转换脚本使用项目级缓存存储依赖检测结果（位置：`.claude/cache/docugenius-converter/`）。

- **缓存策略**：永不过期，仅在 `requirements.txt` 变更时自动失效
- **清除缓存**：如需强制重新检测依赖，运行 `./convert.sh --clear-cache`


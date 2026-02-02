---
name: docugenius-converter
description: 双向文档转换工具。将 Word (.docx)、Excel (.xlsx)、PowerPoint (.pptx) 和 PDF (.pdf) 转换为 Markdown 格式；将 Markdown (.md) 转换为 Word (.docx) 格式。当用户请求转换文档、需要分析文档内容、或批量处理文档时使用。
---

# DocuGenius Document Converter

双向文档转换工具：
- **Office/PDF → Markdown**：将 Office 文档和 PDF 转换为 AI 友好的 Markdown 格式
- **Markdown → Word**：将 Markdown 文件转换为专业的 Word 文档

## 环境要求

### Python 环境（必需）

此 skill 依赖 Python 3.6+，用于 Office/PDF 转 Markdown 功能。

### Node.js 环境（可选）

Markdown 转 Word 功能需要 Node.js 环境。如果不需要此功能，可以跳过 Node.js 安装。

## Python 环境检测与引导（Claude 操作指南）

**重要**：在执行转换前，你必须先检测用户的 Python 环境。

### 自动检测流程

使用 Bash 工具检测 Python 环境：

```bash
# Windows 用户
python --version

# macOS/Linux 用户
python3 --version
```

### 处理不同情况

**情况 1：Python 未安装**

如果命令失败（command not found 或类似错误），告知用户：

```
您的系统未安装 Python。此工具需要 Python 3.6 或更高版本。

请按照以下步骤安装：

Windows:
1. 访问 https://www.python.org/downloads/
2. 下载并运行安装程序
3. 重要：勾选 "Add Python to PATH" 选项
4. 安装完成后重启终端

macOS:
brew install python3

Linux (Ubuntu/Debian):
sudo apt update && sudo apt install python3 python3-pip

安装完成后，请告诉我，我会继续帮您完成文档转换。
```

**情况 2：Python 版本过低**

检测版本是否 >= 3.6：

```bash
python -c "import sys; print(sys.version_info >= (3, 6))"
```

如果返回 False，告知用户需要升级。

**情况 3：Python 已安装且版本符合**

继续执行依赖安装流程：

1. 引导用户运行安装脚本：
   - Windows: `install.bat`
   - macOS/Linux: `./install.sh`

2. 或使用 Bash 工具直接安装：
   ```bash
   python -m pip install -r requirements.txt
   ```

3. 如果安装失败，检查错误信息并引导用户解决

### 依赖库缺失处理

如果转换脚本报错"缺少依赖库"，使用 Bash 工具安装对应的库：

```bash
# 按需安装（推荐）
pip install python-docx  # Word 文档
pip install openpyxl     # Excel 文件
pip install python-pptx  # PowerPoint 文件
pip install pdfplumber   # PDF 文件

# 或安装全部
pip install -r requirements.txt
```

## 快速开始

### 安装依赖

**Windows:**
```bash
install.bat
```

**macOS/Linux:**
```bash
chmod +x install.sh
./install.sh
```

**手动安装:**
```bash
pip install -r requirements.txt
```

### 转换单个文档

```python
from scripts.convert_document import convert_document

result = convert_document(
    file_path='/path/to/document.docx'
)

if result['success']:
    print(f"转换成功: {result['output_path']}")
    print(f"内容预览: {result['markdown_content'][:200]}...")
else:
    print(f"错误: {result['error']}")
```

## 核心工作流程

### 1. 检测文档转换需求

在以下情况触发此 skill：
- 用户明确提到"convert"、"转换"、"transform"、"导出为Markdown"等词汇
- 用户要求分析/读取/理解/总结 .docx/.xlsx/.pptx/.pdf 文件内容
- 用户上传文档文件并询问"这是什么"、"帮我看看"、"分析一下"等
- 用户请求批量处理多个文档
- 用户需要将文档内容用于后续的AI处理

### 2. 检查 Python 环境（自动化）

**你必须使用 Bash 工具自动检测 Python 环境**，不要让用户手动检查。

**步骤 1：检测 Python 是否存在**

```bash
# 根据用户操作系统选择命令
# Windows: python --version
# macOS/Linux: python3 --version
```

**步骤 2：根据检测结果采取行动**

- **如果命令成功**：继续步骤 3 验证文件
- **如果命令失败**：告知用户并提供安装指引（参考"Python 环境检测与引导"章节）

**步骤 3：处理依赖库缺失**

如果转换时报错"缺少依赖库"，使用 Bash 工具安装：

```bash
# 推荐：引导用户运行安装脚本
# Windows: install.bat
# macOS/Linux: ./install.sh

# 或直接安装
python -m pip install -r requirements.txt
```

### 3. 验证文件

检查文件是否存在且格式受支持：
- **Office/PDF 转 Markdown**: `.docx`, `.xlsx`, `.pptx`, `.pdf`
- **Markdown 转 Word**: `.md`
- **不支持**: `.doc`, `.xls`, `.ppt`（旧格式 - 需先转换）

### 4. 转换文档

使用转换脚本：

```bash
python scripts/convert_document.py <file_path> [extract_images] [output_dir]
```

参数：
- `file_path`: 文档路径（必需）
- `extract_images`: `true` 或 `false`（默认: `true`，仅用于 Office/PDF 转换）
- `output_dir`: 可选的输出目录
  - Office/PDF → Markdown: 默认为 `Markdown/` 子目录
  - Markdown → Word: 默认为 `Word/` 子目录

**注意**：Markdown 转 Word 功能需要 Node.js 环境。如果未安装 Node.js，转换时会提示错误。

### 5. 处理结果

脚本输出 JSON 格式结果：
- `success`: 布尔值，表示转换是否成功
- `markdown_content`: 转换后的 Markdown 文本
- `output_path`: 保存的 .md 文件路径
- `error`: 错误信息（如果转换失败）

### 6. 呈现给用户

转换成功后：
1. 显示输出路径
2. 可选地显示内容预览
3. 使用转换后的内容继续处理用户的原始请求

## 常见模式

### 模式 1: Office/PDF 转 Markdown

用户："把这个 Word 文档转换成 Markdown"

```python
result = convert_document('/path/to/report.docx')
if result['success']:
    # 读取并使用 markdown 内容
    with open(result['output_path'], 'r', encoding='utf-8') as f:
        content = f.read()
    # 现在可以分析或处理内容
```

### 模式 2: Markdown 转 Word

用户："把这个 Markdown 文件转换成 Word 文档"

```python
result = convert_document('/path/to/document.md')
if result['success']:
    print(f"转换成功: {result['output_path']}")
else:
    # 检查是否是 Node.js 环境问题
    if 'Node.js' in result['error']:
        print("需要安装 Node.js 才能使用 Markdown 转 Word 功能")
```

### 模式 3: 分析文档内容

用户："把这个 Word 文档转换成 Markdown"

```python
result = convert_document('/path/to/report.docx')
if result['success']:
    # 读取并使用 markdown 内容
    with open(result['output_path'], 'r', encoding='utf-8') as f:
        content = f.read()
    # 现在可以分析或处理内容
```

### 模式 3: 分析文档内容

用户："这个 PDF 的主要内容是什么？"

```python
# 首先转换
result = convert_document('/path/to/document.pdf')
if result['success']:
    # 直接使用 markdown 内容
    content = result['markdown_content']
    # 分析和总结
    # ... 你的分析逻辑 ...
```

### 模式 4: 批量处理

用户："这个 PDF 的主要内容是什么？"

```python
# 首先转换
result = convert_document('/path/to/document.pdf')
if result['success']:
    # 直接使用 markdown 内容
    content = result['markdown_content']
    # 分析和总结
    # ... 你的分析逻辑 ...
```

### 模式 4: 批量处理

用户："转换这个文件夹里的所有文档"

```python
import os
from pathlib import Path

folder = '/path/to/documents'
supported_exts = ['.docx', '.xlsx', '.pptx', '.pdf', '.md']

for file in Path(folder).rglob('*'):
    if file.suffix.lower() in supported_exts:
        result = convert_document(str(file))
        if result['success']:
            print(f"✓ 已转换: {file.name}")
        else:
            print(f"✗ 失败: {file.name} - {result['error']}")
```

或使用批量转换模式：

```bash
python scripts/convert_document.py --batch /path/to/documents
```

## 依赖库

### Python 依赖（Office/PDF 转 Markdown）

- **python-docx**: 处理 Word 文档
- **openpyxl**: 处理 Excel 文件
- **python-pptx**: 处理 PowerPoint 文件
- **pdfplumber**: 处理 PDF 文件

所有依赖都很轻量，总大小约 10-15MB。

### Node.js 依赖（Markdown 转 Word）

- **docx**: DOCX 文档生成库
- **jsdom**: 提供 DOM 环境支持

总大小约 20-30MB。

## 支持的格式

详细的格式支持信息请参见 [supported-formats.md](references/supported-formats.md)。

快速参考：

**Office/PDF → Markdown:**
- **Word (.docx)**: 文本、标题（Heading 1-6）、列表、表格、格式（粗体/斜体） - 质量优秀
- **Excel (.xlsx)**: 表格、多工作表 - 质量优秀
- **PowerPoint (.pptx)**: 幻灯片文本 - 质量良好
- **PDF (.pdf)**: 文本、表格 - 质量取决于 PDF 类型

**Markdown → Word:**
- **Markdown (.md)**: 标题（H1-H6）、列表（有序/无序）、表格、代码块、粗体/斜体、引用块、链接、图片占位符 - 质量优秀

## 安全性特性

此 skill 包含多项安全性改进，确保稳定可靠的转换：

### 文件大小保护
- **硬性限制**：100MB（超过此大小将拒绝转换）
- **推荐大小**：< 50MB（更大的文件转换时间较长）
- **自动检测**：在转换前自动检查文件大小，防止内存溢出

### 详细错误分类
转换失败时提供精确的错误信息：
- **权限错误**：无法读取文件或写入输出目录
- **内存错误**：文件过大导致内存不足
- **文件未找到**：文件路径不存在
- **系统错误**：磁盘空间不足等系统级问题
- **转换错误**：文件损坏或格式不兼容

### 数据安全处理
- **空数据处理**：安全处理空表格、空单元格等边界情况
- **特殊字符转义**：自动转义 Markdown 表格分隔符（|）
- **编码保护**：Windows 平台自动使用 UTF-8 编码

### 按需依赖检查
- 只检查当前文件类型所需的库
- 避免无关依赖阻塞转换流程
- 提供清晰的依赖安装指引

## 转换效果示例

### Word 文档转换效果

**原始 Word 文档内容**：
```
Heading 1: 项目报告
Heading 2: 背景介绍
正文：这是一段**重要**的内容，需要*特别注意*。

Heading 2: 主要发现
- 发现一
- 发现二
- 发现三
```

**转换后的 Markdown**：
```markdown
# 项目报告

## 背景介绍

这是一段**重要**的内容，需要*特别注意*。

## 主要发现

- 发现一
- 发现二
- 发现三
```

### Excel 表格转换效果

**原始 Excel 表格**：
```
姓名    年龄    城市
张三    25     北京
李四    30     上海
```

**转换后的 Markdown**：
```markdown
| 姓名 | 年龄 | 城市 |
| --- | --- | --- |
| 张三 | 25 | 北京 |
| 李四 | 30 | 上海 |
```

### PDF 表格提取效果

**原始 PDF（包含表格）**：
```
产品    价格    库存
笔记本  5000   10
鼠标    50     100
```

**转换后的 Markdown**：
```markdown
## Page 1

| 产品 | 价格 | 库存 |
| --- | --- | --- |
| 笔记本 | 5000 | 10 |
| 鼠标 | 50 | 100 |

### 文本内容

其他文本内容...
```

## 错误处理

常见错误及解决方案：

### 依赖相关错误

**"缺少依赖库: python-docx"**
- **原因**：未安装所需的 Python 库
- **解决方案**：运行 `install.bat`（Windows）或 `install.sh`（macOS/Linux）
- **手动安装**：`pip install python-docx openpyxl python-pptx pdfplumber`
- **按需安装**：只需安装当前文件类型所需的库（例如只转换 Word 文档时只需 `pip install python-docx`）

### 文件相关错误

**"文件不存在: /path/to/file.docx"**
- **原因**：文件路径不正确或文件已被删除
- **解决方案**：
  - 验证文件路径是否正确
  - 使用绝对路径避免混淆
  - 检查文件是否存在

**"文件过大: 150.00MB，超过限制 100MB"**
- **原因**：文件超过 100MB 的安全限制
- **解决方案**：
  - 压缩文件内容（删除不必要的图片、数据）
  - 分割为多个较小的文件
  - 对于 PDF，考虑分页处理

**"不支持的文件格式: .doc"**
- **原因**：使用了旧版 Office 格式
- **解决方案**：
  - 使用 Microsoft Office 或 LibreOffice 将文件另存为新格式
  - .doc → .docx
  - .xls → .xlsx
  - .ppt → .pptx

### 权限相关错误

**"权限不足: 无法读取文件或写入输出目录"**
- **原因**：没有文件读取权限或输出目录写入权限
- **解决方案**：
  - 检查文件权限，确保有读取权限
  - 检查输出目录权限，确保有写入权限
  - Windows：右键文件 → 属性 → 安全
  - macOS/Linux：使用 `chmod` 命令修改权限

### 系统资源错误

**"内存不足: 文件可能过大，请尝试处理较小的文件"**
- **原因**：文件过大导致内存不足
- **解决方案**：
  - 关闭其他程序释放内存
  - 处理较小的文件
  - 增加系统虚拟内存

**"系统错误: [Errno 28] No space left on device"**
- **原因**：磁盘空间不足
- **解决方案**：
  - 清理磁盘空间
  - 更改输出目录到有足够空间的磁盘

### 转换质量问题

**"转换成功但内容不完整"**
- **可能原因**：
  - 文件损坏或格式不标准
  - 使用了不支持的特殊格式
- **解决方案**：
  - 用 Office 软件打开文件验证是否正常
  - 检查是否使用了嵌入对象、特殊字体等
  - 尝试将文件另存为新文件后再转换

**"PDF 表格提取不准确"**
- **可能原因**：
  - PDF 表格结构复杂
  - 扫描型 PDF 没有表格结构信息
- **解决方案**：
  - 检查 PDF 是否为文本型（可以选中文字）
  - 扫描型 PDF 建议先进行 OCR 处理
  - 复杂表格可能需要手动调整输出结果

## 最佳实践

1. **使用新版 Office 格式**: .docx, .xlsx, .pptx（不是 .doc, .xls, .ppt）
2. **文件大小**: 建议 < 50MB，超过 100MB 将被拒绝
3. **使用标准样式**: Word 文档使用内置的 Heading 1-6 样式以获得最佳标题识别效果
4. **列表格式**: 使用 Word 的内置列表功能，而不是手动添加符号
5. **表格设计**: 简单表格转换效果最好，避免复杂的嵌套和合并单元格
6. **PDF 类型**: 文本型 PDF 效果最好，扫描型 PDF 建议先进行 OCR 处理
7. **检查输出质量**: 转换后建议检查输出质量，必要时手动调整
8. **批量处理**: 使用批量转换模式处理多个文件，提高效率

## 完整示例（Claude Code 工作流程）

### 示例 1：用户请求分析文档（Python 已安装）

```
用户："分析这个 Excel 文件中的数据：/path/to/data.xlsx"

Claude 操作流程：

1. 检测到文档分析需求，触发此 skill

2. 使用 Bash 工具检查 Python 环境：
   bash: python --version
   输出: Python 3.9.0

3. Python 已安装，直接转换文档：
   bash: python scripts/convert_document.py /path/to/data.xlsx

4. 转换成功，读取 Markdown 内容并分析

5. 向用户展示分析结果
```

### 示例 2：用户请求转换文档（Python 未安装）

```
用户："把这个 Word 文档转换成 Markdown：report.docx"

Claude 操作流程：

1. 检测到文档转换需求，触发此 skill

2. 使用 Bash 工具检查 Python 环境：
   bash: python --version
   错误: 'python' 不是内部或外部命令

3. Python 未安装，告知用户：

   "您的系统未安装 Python。此工具需要 Python 3.6 或更高版本。

   请按照以下步骤安装：

   Windows:
   1. 访问 https://www.python.org/downloads/
   2. 下载并运行安装程序
   3. 重要：勾选 'Add Python to PATH' 选项
   4. 安装完成后重启终端

   安装完成后，请告诉我，我会继续帮您完成文档转换。"

4. 等待用户安装 Python 后继续
```

### 示例 3：Python 已安装但依赖库缺失

```
用户："转换这个 PDF：document.pdf"

Claude 操作流程：

1. 检测到文档转换需求，触发此 skill

2. 使用 Bash 工具检查 Python 环境：
   bash: python --version
   输出: Python 3.9.0

3. Python 已安装，尝试转换：
   bash: python scripts/convert_document.py document.pdf
   错误: 缺少依赖库: pdfplumber

4. 依赖库缺失，使用 Bash 工具安装：
   bash: python -m pip install pdfplumber

5. 安装成功，重新转换文档

6. 转换成功，向用户展示结果
```

## 系统要求

- **Python**: 3.6 或更高版本（必需）
- **Node.js**: 14.0 或更高版本（可选，仅用于 Markdown 转 Word）
- **操作系统**: Windows、macOS、Linux
- **磁盘空间**: 约 50MB（包括所有依赖）

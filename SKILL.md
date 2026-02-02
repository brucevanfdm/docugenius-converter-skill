---
name: docugenius-converter
description: 文档转换工具，将 Word (.docx)、Excel (.xlsx)、PowerPoint (.pptx) 和 PDF (.pdf) 转换为 Markdown 格式。当用户请求转换文档、需要分析文档内容（"分析这个 Word"、"读取这个 PDF"）、或批量处理文档时使用。
---

# DocuGenius Document Converter

独立的文档转换工具，将 Office 文档和 PDF 转换为 AI 友好的 Markdown 格式。

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

### 2. 验证文件

检查文件是否存在且格式受支持：
- **支持**: `.docx`, `.xlsx`, `.pptx`, `.pdf`
- **不支持**: `.doc`, `.xls`, `.ppt`（旧格式 - 需先转换）

### 3. 转换文档

使用转换脚本：

```bash
python scripts/convert_document.py <file_path> [extract_images] [output_dir]
```

参数：
- `file_path`: 文档路径（必需）
- `extract_images`: `true` 或 `false`（默认: `true`，**注意：当前版本此参数保留供未来使用，暂不影响转换结果**）
- `output_dir`: 可选的输出目录（默认: 同目录下的 `Markdown/` 子目录）

### 4. 处理结果

脚本输出 JSON 格式结果：
- `success`: 布尔值，表示转换是否成功
- `markdown_content`: 转换后的 Markdown 文本
- `output_path`: 保存的 .md 文件路径
- `error`: 错误信息（如果转换失败）

### 5. 呈现给用户

转换成功后：
1. 显示输出路径
2. 可选地显示内容预览
3. 使用转换后的内容继续处理用户的原始请求

## 常见模式

### 模式 1: 单文件转换

用户："把这个 Word 文档转换成 Markdown"

```python
result = convert_document('/path/to/report.docx')
if result['success']:
    # 读取并使用 markdown 内容
    with open(result['output_path'], 'r', encoding='utf-8') as f:
        content = f.read()
    # 现在可以分析或处理内容
```

### 模式 2: 分析文档内容

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

### 模式 3: 批量处理

用户："转换这个文件夹里的所有文档"

```python
import os
from pathlib import Path

folder = '/path/to/documents'
supported_exts = ['.docx', '.xlsx', '.pptx', '.pdf']

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

此 skill 使用以下轻量级 Python 库：

- **python-docx**: 处理 Word 文档
- **openpyxl**: 处理 Excel 文件
- **python-pptx**: 处理 PowerPoint 文件
- **pdfplumber**: 处理 PDF 文件

所有依赖都很轻量，总大小约 10-15MB。

## 支持的格式

详细的格式支持信息请参见 [supported-formats.md](references/supported-formats.md)。

快速参考：
- **Word (.docx)**: 文本、标题（Heading 1-6）、列表、表格、格式（粗体/斜体） - 质量优秀
- **Excel (.xlsx)**: 表格、多工作表 - 质量优秀
- **PowerPoint (.pptx)**: 幻灯片文本 - 质量良好
- **PDF (.pdf)**: 文本、表格 - 质量取决于 PDF 类型

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

## 完整示例

```python
# 用户问："分析这个 Excel 文件中的数据"

# 步骤 1: 转换 Excel 文件
result = convert_document(
    file_path='/path/to/data.xlsx'
)

# 步骤 2: 检查转换是否成功
if not result['success']:
    print(f"转换失败: {result['error']}")
    # 询问用户检查文件或安装依赖
    exit(1)

# 步骤 3: 使用转换后的内容
markdown_content = result['markdown_content']

# 步骤 4: 分析数据
# Excel 表格现在是 Markdown 表格格式
# 你可以解析和分析它们

# 步骤 5: 通知用户
print(f"已将 Excel 文件转换为: {result['output_path']}")
print("分析结果:")
# ... 显示你的分析 ...
```

## 系统要求

- **Python**: 3.6 或更高版本
- **操作系统**: Windows、macOS、Linux
- **磁盘空间**: 约 20MB（包括依赖）

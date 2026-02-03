---
name: docugenius-converter
description: 双向文档转换工具，将 Word (.docx)、Excel (.xlsx)、PowerPoint (.pptx) 和 PDF (.pdf) 转换为 AI 友好的 Markdown 格式，或将 Markdown (.md) 转换为 Word (.docx) 格式。当用户请求以下操作时使用：(1) 明确请求文档转换，包括任何包含"转换"、"转为"、"转成"、"convert"、"导出"、"export"等词汇的请求（例如："转换文档"、"把这个文件转为docx"、"convert to markdown"、"导出为Word"）；(2) 需要 AI 理解文档内容（"帮我分析这个 Word 文件"、"读取这个 PDF"、"总结这个 Excel"）；(3) 上传文档文件并询问内容（"这是什么"、"帮我看看"）；(4) 任何涉及 .docx、.xlsx、.pptx、.pdf、.md 文件格式转换的请求。
---
# DocuGenius Document Converter

双向文档转换，自动处理依赖安装和缓存。

## Quick Reference

| 操作                   | 命令                           | 输出位置             |
| ---------------------- | ------------------------------ | -------------------- |
| Office/PDF → Markdown | `./convert.sh <file>`        | 同目录 `Markdown/` |
| Markdown → Word       | `./convert.sh <file.md>`     | 同目录 `Word/`     |
| 批量转换               | `./convert.sh --batch <dir>` | 同上                 |

## 工作流程

```
用户请求转换 → 直接运行 ./convert.sh → 解析 JSON 输出 → 处理结果
```

**关键原则**：
1. **不要预先检查任何依赖**（Python 库、Node.js 等）
2. 直接执行转换命令
3. 只在转换失败（`success: false`）时才根据错误信息处理

## 执行命令

```bash
# 单文件转换（依赖自动安装）
./convert.sh /path/to/document.docx

# 自定义输出目录
./convert.sh /path/to/file.pdf true /custom/output

# 批量转换
./convert.sh --batch /path/to/documents
```

**Windows 用户**：使用 `convert.bat` 替代 `./convert.sh`

## 解析输出

脚本返回 JSON，关键字段：

```json
{
  "success": true,
  "output_path": "/path/to/output.md",
  "markdown_content": "# 转换后的内容..."
}
```

- `success`: 转换是否成功
- `output_path`: 输出文件路径
- `markdown_content`: Markdown 内容（方便直接分析）
- `error`: 错误信息（失败时）

## 错误处理

**仅在转换失败时（返回 `success: false`）才处理错误**：

| 错误类型                      | 处理方法                                                |
| ----------------------------- | ------------------------------------------------------- |
| Python 依赖缺失              | 脚本会自动安装，如失败则运行 `pip install --user xxx` |
| `未找到 Node.js`             | 仅在 MD→DOCX 转换失败且报此错误时，才提示安装 Node.js  |
| `Node.js 依赖未安装`         | 在 `scripts/md_to_docx` 目录运行 `npm install`        |
| `文件不存在`                 | 提示用户验证文件路径                                    |
| `不支持的文件格式: .doc`     | 提示用户先转换为 .docx                                  |
| `文件过大`                   | 提示超过 100MB 限制                                      |

## 支持的格式

| 格式  | 转换方向 | 质量            |
| ----- | -------- | --------------- |
| .docx | ↔       | 优秀            |
| .xlsx | →       | 优秀            |
| .pptx | →       | 良好            |
| .pdf  | →       | 取决于 PDF 类型 |
| .md   | ↔       | 优秀            |

## 注意事项

**重要：**
- **绝对不要在执行转换前检查任何依赖**（包括 Python、Node.js、npm 包等）
- 直接执行转换命令，让脚本自己检测和处理依赖
- 只在转换失败时才根据返回的错误信息采取行动

其他：
- 依赖使用项目级缓存，首次检测后无需重复检查
- Python 依赖会自动安装到用户目录，无需虚拟环境
- .doc/.xls/.ppt 旧格式需先转换为对应的新格式

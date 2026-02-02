---
name: docugenius-converter
description: Use when user requests document conversion, export, or AI analysis of Office/PDF files. Trigger words: "转换", "convert", "导出", "export", "分析这个文件", or any .docx/.xlsx/.pptx/.pdf/.md file operation.
---
# DocuGenius Document Converter

双向文档转换，自动处理依赖安装和缓存。

## Quick Reference

| 操作 | 命令 | 输出位置 |
|------|------|----------|
| Office/PDF → Markdown | `./convert.sh <file>` | 同目录 `Markdown/` |
| Markdown → Word | `./convert.sh <file.md>` | 同目录 `Word/` |
| 批量转换 | `./convert.sh --batch <dir>` | 同上 |

## 工作流程

```
用户请求转换 → 直接运行 ./convert.sh → 解析 JSON 输出 → 处理结果
```

**关键原则**：依赖会自动安装，无需预检查。转换失败时再处理错误。

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

仅在转换失败时处理：

| 错误 | 解决方案 |
|------|----------|
| `缺少依赖库: xxx` | 脚本会自动安装，如失败则手动 `pip install --user xxx` |
| `文件不存在` | 验证路径，使用绝对路径 |
| `不支持的文件格式: .doc` | 提示用户先转换为 .docx |
| `未找到 Node.js` | 仅 Markdown→Word 需要，提示安装 Node.js |
| `文件过大` | 超过 100MB 限制 |

## 支持的格式

| 格式 | 转换方向 | 质量 |
|------|----------|------|
| .docx | ↔ | 优秀 |
| .xlsx | → | 优秀 |
| .pptx | → | 良好 |
| .pdf | → | 取决于 PDF 类型 |
| .md | ↔ | 优秀 |

## 注意事项

- 依赖使用项目级缓存，首次检测后无需重复检查
- 自动安装到用户目录，无需虚拟环境
- .doc/.xls/.ppt 旧格式需先转换



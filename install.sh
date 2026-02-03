#!/bin/bash
# DocuGenius Converter - 环境检查说明

cat << 'EOF'
========================================
DocuGenius Converter Skill
========================================

本 Skill 不需要预安装依赖！

首次使用时，依赖会自动安装到用户目录：
- 无需虚拟环境
- 不受 macOS PEP 668 限制
- 所有项目共享

----------------------------------------
环境要求
----------------------------------------

1. Python 3.6+（必需）

   检查版本：
     python3 --version

   如未安装：
     macOS:   brew install python3
     Ubuntu:  sudo apt install python3
     其他:    https://www.python.org/downloads/

2. Node.js 14+（可选）

   仅用于 Markdown → Word 转换功能

   检查版本：
     node --version

   如需此功能，请安装 Node.js：
     macOS:   brew install node
     其他:    https://nodejs.org/

----------------------------------------
手动安装（可选）
----------------------------------------

如果自动安装失败，可手动安装依赖：

# Python 依赖
pip install --user python-docx openpyxl python-pptx pdfplumber

# Node.js 依赖（可选，仅用于 MD → DOCX）
# 默认会自动安装到用户级共享目录：
#   macOS/Linux: ~/.docugenius/node/md_to_docx
#   Windows:     %LOCALAPPDATA%/DocuGenius/node/md_to_docx
# 可通过环境变量 DOCUGENIUS_NODE_HOME 指定共享目录
#
# 方案 A：本地安装（当前项目）
cd scripts/md_to_docx && npm install
# 方案 B：共享安装（用户目录）
# cd ~/.docugenius/node/md_to_docx && npm install

----------------------------------------
使用方法
----------------------------------------

在 Claude Code 中直接对话：

  "帮我分析这个 report.docx"
  "把这个 document.pdf 转成 markdown"
  "把 notes.md 导出为 Word 文档"

EOF

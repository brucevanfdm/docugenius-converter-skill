#!/bin/bash
# 安装 DocuGenius Converter 依赖

echo "正在安装 DocuGenius Converter 依赖..."

# 检查 Python 是否安装
if ! command -v python3 &> /dev/null && ! command -v python &> /dev/null; then
    echo "错误: 未找到 Python。请先安装 Python 3.6 或更高版本。"
    exit 1
fi

# 使用 python3 或 python
if command -v python3 &> /dev/null; then
    PYTHON_CMD=python3
else
    PYTHON_CMD=python
fi

echo "使用 Python: $PYTHON_CMD"

# 安装依赖
$PYTHON_CMD -m pip install -r requirements.txt

if [ $? -eq 0 ]; then
    echo "✓ 依赖安装成功！"
    echo ""
    echo "现在可以使用转换脚本了："
    echo "  $PYTHON_CMD scripts/convert_document.py <file_path>"
else
    echo "✗ 依赖安装失败。请检查错误信息。"
    exit 1
fi

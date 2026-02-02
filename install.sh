#!/bin/bash
# 安装 DocuGenius Converter 依赖

echo "正在安装 DocuGenius Converter 依赖..."
echo ""

# 检查 Python 是否安装
if ! command -v python3 &> /dev/null && ! command -v python &> /dev/null; then
    echo "错误: 未找到 Python。请先安装 Python 3.6 或更高版本。"
    echo ""
    echo "安装方法："
    echo "  macOS:   brew install python3"
    echo "           或访问 https://www.python.org/downloads/macos/"
    echo ""
    echo "  Ubuntu:  sudo apt update && sudo apt install python3 python3-pip"
    echo "  CentOS:  sudo yum install python3 python3-pip"
    echo ""
    echo "  其他系统: https://www.python.org/downloads/"
    exit 1
fi

# 使用 python3 或 python
if command -v python3 &> /dev/null; then
    PYTHON_CMD=python3
else
    PYTHON_CMD=python
fi

echo "检测到 Python: $PYTHON_CMD"
$PYTHON_CMD --version
echo ""

# 检查 Python 版本是否 >= 3.6
$PYTHON_CMD -c "import sys; exit(0 if sys.version_info >= (3, 6) else 1)" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "错误: Python 版本过低。需要 Python 3.6 或更高版本。"
    echo "当前版本: $($PYTHON_CMD --version)"
    echo ""
    echo "请升级 Python:"
    echo "  macOS:   brew upgrade python3"
    echo "  Ubuntu:  sudo apt install python3.8"
    echo "  其他:    https://www.python.org/downloads/"
    exit 1
fi

echo "✓ Python 版本符合要求"
echo ""

# 安装依赖
echo "正在安装依赖库..."
$PYTHON_CMD -m pip install -r requirements.txt

if [ $? -eq 0 ]; then
    echo ""
    echo "✓ 依赖安装成功！"
    echo ""
    echo "现在可以使用转换脚本了："
    echo "  $PYTHON_CMD scripts/convert_document.py <file_path>"
else
    echo ""
    echo "✗ 依赖安装失败。请检查错误信息。"
    exit 1
fi

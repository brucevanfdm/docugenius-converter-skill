#!/bin/bash
# 安装 DocuGenius Converter 依赖

echo "正在安装 DocuGenius Converter 依赖..."
echo ""

# ========== Python 环境检查 ==========
echo "[1/2] 检查 Python 环境..."

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

echo "[OK] Python 版本符合要求"
echo ""

# ========== 创建全局共享虚拟环境 ==========
echo "正在创建全局共享虚拟环境..."

# 全局虚拟环境位置（所有项目共享）
GLOBAL_VENV_BASE="$HOME/.claude/venvs"
VENV_DIR="$GLOBAL_VENV_BASE/docugenius-converter"

# 创建全局 venv 基础目录
mkdir -p "$GLOBAL_VENV_BASE"

# 如果虚拟环境已存在，先删除
if [ -d "$VENV_DIR" ]; then
    echo "检测到已存在的虚拟环境，正在删除..."
    rm -rf "$VENV_DIR"
fi

# 创建虚拟环境
$PYTHON_CMD -m venv "$VENV_DIR"

if [ $? -ne 0 ]; then
    echo ""
    echo "[ERROR] 虚拟环境创建失败。请确保已安装 python3-venv。"
    echo "  macOS:   brew install python@3.11 或 python@3.12"
    echo "  Ubuntu:  sudo apt install python3-venv"
    exit 1
fi

echo "[OK] 全局虚拟环境创建成功: $VENV_DIR"
echo "[说明] 所有项目共享此虚拟环境，无需重复安装依赖"
echo ""

# 激活虚拟环境并安装依赖
echo "正在安装 Python 依赖库..."
source "$VENV_DIR/bin/activate"
pip install --upgrade pip >/dev/null 2>&1
pip install -r requirements.txt

if [ $? -ne 0 ]; then
    echo ""
    echo "[ERROR] Python 依赖安装失败。请检查错误信息。"
    exit 1
fi

deactivate

echo "[OK] Python 依赖安装成功！"
echo ""

# ========== Node.js 环境检查（用于 Markdown 转 DOCX）==========
echo "[2/2] 检查 Node.js 环境（用于 Markdown 转 DOCX）..."

if ! command -v node &> /dev/null; then
    echo ""
    echo "警告: 未找到 Node.js。Markdown 转 DOCX 功能将不可用。"
    echo "如需此功能，请安装 Node.js:"
    echo "  macOS:   brew install node"
    echo "  Ubuntu:  sudo apt install nodejs npm"
    echo "  其他:    https://nodejs.org/"
    echo ""
    echo "其他功能（Office/PDF 转 Markdown）可正常使用。"
else
    echo "检测到 Node.js:"
    node --version
    echo ""

    # 安装 Node.js 依赖
    echo "正在安装 Node.js 依赖..."
    SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
    cd "$SCRIPT_DIR/scripts/md_to_docx"
    npm install

    if [ $? -ne 0 ]; then
        echo ""
        echo "[ERROR] Node.js 依赖安装失败。Markdown 转 DOCX 功能将不可用。"
        echo "请手动在 scripts/md_to_docx 目录下运行: npm install"
    else
        echo "[OK] Node.js 依赖安装成功！"
    fi

    cd "$SCRIPT_DIR"
fi

echo ""
echo "========================================"
echo "安装完成！"
echo "========================================"
echo ""
echo "支持的转换:"
echo "  - Office/PDF 转 Markdown: .docx, .xlsx, .pptx, .pdf"
echo "  - Markdown 转 Word: .md (需要 Node.js)"
echo ""
echo "使用方法:"
echo "  ./convert.sh <file_path>"
echo ""
echo "或者手动激活全局虚拟环境:"
echo "  source ~/.claude/venvs/docugenius-converter/bin/activate"
echo "  python scripts/convert_document.py <file_path>"
echo ""

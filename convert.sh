#!/bin/bash
# DocuGenius Converter - 使用共享虚拟环境运行转换脚本

# 全局虚拟环境位置（所有 Claude Code skills 共享基础目录）
GLOBAL_VENV_BASE="$HOME/.claude/venvs"
GLOBAL_VENV_DIR="$GLOBAL_VENV_BASE/docugenius-converter"

# 优先使用当前项目的虚拟环境（如果存在）
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
LOCAL_VENV_DIR="$SCRIPT_DIR/.venv"

# 决定使用哪个虚拟环境
if [ -d "$LOCAL_VENV_DIR" ]; then
    VENV_DIR="$LOCAL_VENV_DIR"
elif [ -d "$GLOBAL_VENV_DIR" ]; then
    VENV_DIR="$GLOBAL_VENV_DIR"
else
    echo "错误: 虚拟环境不存在。"
    echo "请先在 skill 目录运行: install.sh"
    echo "  Skill 目录: ~/.claude/skills/docugenius-converter/"
    exit 1
fi

# 激活虚拟环境并运行转换脚本
source "$VENV_DIR/bin/activate"
python "$SCRIPT_DIR/scripts/convert_document.py" "$@"
deactivate

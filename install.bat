@echo off
REM 安装 DocuGenius Converter 依赖 (Windows)

echo 正在安装 DocuGenius Converter 依赖...

REM 检查 Python 是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到 Python。请先安装 Python 3.6 或更高版本。
    echo 下载地址: https://www.python.org/downloads/
    echo.
    echo 安装时请务必勾选 "Add Python to PATH" 选项！
    pause
    exit /b 1
)

echo 检测到 Python:
python --version

REM 检查 Python 版本是否 >= 3.6
python -c "import sys; exit(0 if sys.version_info >= (3, 6) else 1)" >nul 2>&1
if errorlevel 1 (
    echo.
    echo 错误: Python 版本过低。需要 Python 3.6 或更高版本。
    echo 当前版本:
    python --version
    echo.
    echo 请升级 Python: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo ✓ Python 版本符合要求
echo.

REM 安装依赖
echo 正在安装依赖库...
python -m pip install -r requirements.txt

if errorlevel 1 (
    echo.
    echo ✗ 依赖安装失败。请检查错误信息。
    pause
    exit /b 1
)

echo.
echo ✓ 依赖安装成功！
echo.
echo 现在可以使用转换脚本了：
echo   python scripts\convert_document.py ^<file_path^>
echo.
pause

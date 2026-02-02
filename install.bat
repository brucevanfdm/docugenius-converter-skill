@echo off
REM 安装 DocuGenius Converter 依赖 (Windows)

echo 正在安装 DocuGenius Converter 依赖...

REM 检查 Python 是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到 Python。请先安装 Python 3.6 或更高版本。
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo 使用 Python:
python --version

REM 安装依赖
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

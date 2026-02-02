@echo off
REM 安装 DocuGenius Converter 依赖 (Windows)

echo 正在安装 DocuGenius Converter 依赖...
echo.

REM ========== Python 环境检查 ==========
echo [1/2] 检查 Python 环境...

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

echo [OK] Python 版本符合要求
echo.

REM 安装 Python 依赖
echo 正在安装 Python 依赖库...
python -m pip install -r requirements.txt

if errorlevel 1 (
    echo.
    echo × Python 依赖安装失败。请检查错误信息。
    pause
    exit /b 1
)

echo [OK] Python 依赖安装成功！
echo.

REM ========== Node.js 环境检查（用于 Markdown 转 DOCX）==========
echo [2/2] 检查 Node.js 环境（用于 Markdown 转 DOCX）...

node --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo 警告: 未找到 Node.js。Markdown 转 DOCX 功能将不可用。
    echo 如需此功能，请安装 Node.js: https://nodejs.org/
    echo.
    echo 其他功能（Office/PDF 转 Markdown）可正常使用。
    echo.
    goto :done
)

echo 检测到 Node.js:
node --version

REM 安装 Node.js 依赖
echo 正在安装 Node.js 依赖...
cd scripts\md_to_docx
call npm install

if errorlevel 1 (
    echo.
    echo [ERROR] Node.js 依赖安装失败。Markdown 转 DOCX 功能将不可用。
    echo 请手动在 scripts\md_to_docx 目录下运行: npm install
    cd ..\..
    goto :done
)

cd ..\..
echo [OK] Node.js 依赖安装成功！
echo.

:done
echo.
echo ========================================
echo 安装完成！
echo ========================================
echo.
echo 支持的转换:
echo   - Office/PDF 转 Markdown: .docx, .xlsx, .pptx, .pdf
echo   - Markdown 转 Word: .md (需要 Node.js)
echo.
echo 使用方法:
echo   python scripts\convert_document.py ^<file_path^>
echo.
pause

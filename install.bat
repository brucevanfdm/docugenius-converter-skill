@echo off
REM DocuGenius Converter - 环境检查说明

echo ========================================
echo DocuGenius Converter Skill
echo ========================================
echo.
echo 本 Skill 不需要预安装依赖！
echo.
echo 首次使用时，依赖会自动安装到用户目录：
echo - 无需虚拟环境
echo - 不受系统保护限制
echo - 所有项目共享
echo.
echo ----------------------------------------
echo 环境要求
echo ----------------------------------------
echo.
echo 1. Python 3.6+（必需）
echo.
echo    检查版本：
echo      python --version
echo.
echo    如未安装：
echo      访问: https://www.python.org/downloads/
echo      安装时请勾选 "Add Python to PATH"
echo.
echo 2. Node.js 14+（可选）
echo.
echo    仅用于 Markdown 转 Word 转换功能
echo.
echo    检查版本：
echo      node --version
echo.
echo    如需此功能，请安装 Node.js：
echo      访问: https://nodejs.org/
echo.
echo ----------------------------------------
echo 手动安装（可选）
echo ----------------------------------------
echo.
echo 如果自动安装失败，可手动安装依赖：
echo.
echo   pip install --user python-docx openpyxl python-pptx pdfplumber
echo.
echo   REM Node.js 依赖（可选，仅用于 MD 转 DOCX）
echo   REM 默认会自动安装到用户级共享目录：
echo   REM   Windows: %%LOCALAPPDATA%%\DocuGenius\node\md_to_docx
echo   REM 可通过环境变量 DOCUGENIUS_NODE_HOME 指定共享目录
echo   REM
echo   REM 方案 A：本地安装（当前项目）
echo   cd scripts\md_to_docx
echo   npm install
echo   REM 方案 B：共享安装（用户目录）
echo   REM cd %%LOCALAPPDATA%%\DocuGenius\node\md_to_docx
echo   REM npm install
echo.
echo ----------------------------------------
echo 使用方法
echo ----------------------------------------
echo.
echo 在 Claude Code 中直接对话：
echo.
echo   "帮我分析这个 report.docx"
echo   "把这个 document.pdf 转成 markdown"
echo   "把 notes.md 导出为 Word 文档"
echo.
pause

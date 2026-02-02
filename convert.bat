@echo off
REM DocuGenius Converter - 使用共享虚拟环境运行转换脚本

setlocal

REM 全局虚拟环境位置
set "GLOBAL_VENV_BASE=%USERPROFILE%\.claude\venvs"
set "GLOBAL_VENV_DIR=%GLOBAL_VENV_BASE%\docugenius-converter"

REM 优先使用当前项目的虚拟环境（如果存在）
set "SCRIPT_DIR=%~dp0"
set "LOCAL_VENV_DIR=%SCRIPT_DIR%.venv"

REM 决定使用哪个虚拟环境
if exist "%LOCAL_VENV_DIR%" (
    set "VENV_DIR=%LOCAL_VENV_DIR%"
) else if exist "%GLOBAL_VENV_DIR%" (
    set "VENV_DIR=%GLOBAL_VENV_DIR%"
) else (
    echo 错误: 虚拟环境不存在。
    echo 请先在 skill 目录运行: install.bat
    echo   Skill 目录: %%USERPROFILE%%\.claude\skills\docugenius-converter\
    exit /b 1
)

REM 激活虚拟环境并运行转换脚本
call "%VENV_DIR%\Scripts\activate.bat"
python "%SCRIPT_DIR%scripts\convert_document.py" %*
deactivate

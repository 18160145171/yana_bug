@echo off
chcp 65001 >nul
title 质量验收与缺陷分析报告生成器

set TOOL_DIR=F:\bug_report_tool
set TOOL_PY=%TOOL_DIR%\generate_bug_report.py

if not exist "%TOOL_PY%" (
    echo [错误] 未找到工具脚本: %TOOL_PY%
    pause
    exit /b 1
)

python --version >nul 2>nul
if errorlevel 1 (
    py --version >nul 2>nul
    if errorlevel 1 (
        echo [错误] 未检测到 Python，请先安装 Python 3.9+ 并加入 PATH。
        pause
        exit /b 1
    ) else (
        py "%TOOL_PY%"
        goto :end
    )
)

python "%TOOL_PY%"

:end
pause

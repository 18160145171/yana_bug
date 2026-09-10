@echo off
chcp 65001 >nul
title 质量验收报告网页版启动器
cd /d %~dp0

echo 正在启动本地网页工具...
echo 地址: http://127.0.0.1:8512

python --version >nul 2>nul
if errorlevel 1 (
    py --version >nul 2>nul
    if errorlevel 1 (
        echo [错误] 未检测到 Python，请先安装 Python 3.9+ 并加入 PATH。
        pause
        exit /b 1
    ) else (
        start "" "http://127.0.0.1:8512"
        py app_web.py
        pause
        exit /b 0
    )
)

python -c "import flask" >nul 2>nul
if errorlevel 1 (
    echo [提示] 首次运行，正在安装依赖...
    python -m pip install -r requirements.txt
)

start "" "http://127.0.0.1:8512"
python app_web.py
pause

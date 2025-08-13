@echo off
chcp 65001
setlocal

REM 检查Python环境
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo 未检测到Python环境，请先安装Python 3.x
    pause
    exit /b 1
)

REM 安装依赖
pip install -r requirements.txt

REM 执行转换
python excel2json.py

if %errorlevel% equ 0 (
    echo 转换成功完成！
    pause
) else (
    echo 转换过程中出错！
    pause
    exit /b 1
)

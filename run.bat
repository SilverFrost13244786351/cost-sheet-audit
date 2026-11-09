@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"
rem 用项目自带的 Python 3.12（已装好 pandas / openpyxl），不依赖系统里的 3.15
"C:\Users\BYH\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe" src\audit.py
echo.
pause

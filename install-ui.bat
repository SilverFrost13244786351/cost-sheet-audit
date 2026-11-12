@echo off
cd /d "%~dp0"
echo Installing streamlit, please wait...
C:\Users\BYH\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe -m pip install streamlit -i https://pypi.tuna.tsinghua.edu.cn/simple
echo.
echo Done. Now double-click start-ui.bat
pause

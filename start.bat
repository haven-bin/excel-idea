@echo off
chcp 65001 > nul
echo ===================================================
echo   Smart Excel Merger 智能表格合并系统启动中...
echo ===================================================
echo.

set PYTHON_EXE=C:\Python314\python.exe
if not exist %PYTHON_EXE% (
    set PYTHON_EXE=python
)

echo 正在检查服务环境...
%PYTHON_EXE% -m uvicorn main:app --host 127.0.0.1 --port 8000 --reload

pause

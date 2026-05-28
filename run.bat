@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Запуск MOEX Report Generator...
echo.
.venv\Scripts\python main.py
pause

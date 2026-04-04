@echo off
chcp 65001 >nul
cd /d "%~dp0"
python owner_check.py
pause

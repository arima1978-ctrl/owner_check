@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo   売上照合チェック Webサーバー起動中...
echo   ブラウザで http://localhost:3006 にアクセスしてください
echo   停止するには このウィンドウを閉じるか Ctrl+C
echo.
python web.py
pause

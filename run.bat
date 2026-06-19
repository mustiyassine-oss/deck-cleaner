@echo off
title Deck Cleaner — PPTX Layout Cleanup
echo.
echo  ================================================
echo    Deck Cleaner — PPTX Layout Cleanup
echo  ================================================
echo.

cd /d "%~dp0"

if exist .venv\Scripts\activate.bat (
    call .venv\Scripts\activate.bat
)

echo  Starting server on http://127.0.0.1:8080
echo  Press Ctrl+C to stop.
echo.
start "" http://127.0.0.1:8080
python -m uvicorn main:app --host 127.0.0.1 --port 8080
pause

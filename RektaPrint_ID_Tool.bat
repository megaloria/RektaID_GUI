@echo off
cd /d "%~dp0"
python id_photo_processor.py
if %errorlevel% neq 0 (
    echo.
    echo Something went wrong. Press any key to close.
    pause >nul
)

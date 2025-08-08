@echo off
echo PowerPoint to Video Converter
echo ============================
echo.

REM Change to the directory where the script is located
cd /d "%~dp0"

REM Run the Python script
python convert_powerpoint_to_video.py

REM Pause so user can see the results
echo.
pause

@echo off
img inserter @ word by yel
:loop
cls
cd /d "%~dp0"
py resizer.py
echo.
echo Script finished. Restarting...
timeout /t 2 >nul
goto loop

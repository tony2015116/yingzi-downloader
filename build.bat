@echo off
chcp 65001 >nul
cd /d "%~dp0"

rem Find Python: "python" first, then the "py" launcher
set "PY=python"
where python >nul 2>nul
if errorlevel 1 (
    where py >nul 2>nul
    if errorlevel 1 (
        echo [ERROR] Python not found. Install Python 3.8+ from https://www.python.org/downloads/
        echo         and tick "Add python.exe to PATH" during installation.
        pause
        exit /b 1
    )
    set "PY=py -3"
)

echo Building with: %PY%
%PY% build.py %*
if errorlevel 1 (
    echo.
    echo [ERROR] Build failed. See the messages above.
    pause
    exit /b 1
)

echo.
echo Done. The exe is in the "dist" folder.
pause

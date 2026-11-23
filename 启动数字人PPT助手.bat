@echo off
chcp 65001 >nul
cd /d "%~dp0"
set PYTHONPATH=
set PYTHONNOUSERSITE=1

if exist "%~dp0runtime\pythonw.exe" (
    start "" "%~dp0runtime\pythonw.exe" "%~dp0src\app.py"
    exit /b 0
)

if exist "%~dp0.venv\Scripts\pythonw.exe" (
    start "" "%~dp0.venv\Scripts\pythonw.exe" "%~dp0src\app.py"
    exit /b 0
)

where pythonw >nul 2>nul
if %errorlevel%==0 (
    python -c "import PIL, edge_tts" >nul 2>nul
    if %errorlevel%==0 (
        start "" pythonw "%~dp0src\app.py"
        exit /b 0
    )
)

call "%~dp0一键安装并启动.bat"

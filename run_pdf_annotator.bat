@echo off
setlocal

set "APP_DIR=%~dp0"
set "PY_EXE=.venv\Scripts\python.exe"
set "APP_URL=http://127.0.0.1:5000/"
pushd "%APP_DIR%" >nul

if not exist "%PY_EXE%" (
    echo Python environment not found. Run setup.bat first.
    pause
    popd >nul
    exit /b 1
)

echo Starting PDF Annotation Studio...

:: 1. Fire off the server blindly to avoid slow PowerShell polling loops. 
:: (If already running, this new instance just fails to bind to port 5000 and silently exits)
powershell -NoProfile -NonInteractive -Command "Start-Process '%PY_EXE%' -ArgumentList '-c \"from app import app; app.run(debug=False, host=''127.0.0.1'', port=5000)\"' -WindowStyle Hidden"

:: 2. Brief static delay to prevent a race condition before Edge loads
timeout /t 2 >nul

:: 3. Launch Edge in chromeless App Mode, forcing the browser to maximize
start msedge.exe --app="%APP_URL%" --start-maximized

popd >nul
endlocal
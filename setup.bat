@echo off
setlocal EnableDelayedExpansion

title PDF Annotation Studio Setup
echo ===================================================
echo   PDF Annotation Studio - Native Compiler Setup
echo ===================================================

:: 1. Check for Python
python --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python is not installed or not in your system PATH.
    echo Please install Python 3.9 or higher and try again.
    pause
    exit /b
)

:: 2. Setup Virtual Environment
echo.
echo [1/6] Configuring isolated virtual environment...
IF NOT EXIST "venv\" (
    python -m venv venv
)
call venv\Scripts\activate

:: 3. Install Dependencies (Including PyInstaller)
echo.
echo [2/6] Installing Python dependencies...
pip install flask PyMuPDF openpyxl pywebview pyinstaller --quiet

:: 4. Ensure icon.ico exists
echo.
echo [3/6] Verifying application icon...
IF NOT EXIST "icon.ico" (
    echo Generating placeholder icon.ico...
    python -c "import base64; open('icon.ico','wb').write(base64.b64decode('AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAQAABILAAASCwAAAAAAAAAAAAAAAAD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A'))"
)

:: 5. Compile the Application
echo.
echo [4/6] Compiling native Windows executable (This may take a minute)...
:: The icon is permanently embedded into the resulting .exe here
pyinstaller --noconfirm --onedir --windowed --icon="icon.ico" --name="PDF Annotation Studio" --add-data="index.html;." --add-data="icon.ico;." app.py

:: 6. Create Start Menu Folder and Shortcuts
echo.
echo [5/6] Creating Start Menu shortcuts...

set "START_MENU_DIR=%APPDATA%\Microsoft\Windows\Start Menu\Programs\PDF Annotation Studio"
IF NOT EXIST "%START_MENU_DIR%" mkdir "%START_MENU_DIR%"

set "APP_SHORTCUT=%START_MENU_DIR%\PDF Annotation Studio.lnk"
set "UNINSTALL_SHORTCUT=%START_MENU_DIR%\Uninstall.lnk"

:: Point to the newly compiled native executable
set "TARGET_EXE=%CD%\dist\PDF Annotation Studio\PDF Annotation Studio.exe"
set "WORKING_DIR=%CD%\dist\PDF Annotation Studio"

:: Create App Shortcut
powershell -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%APP_SHORTCUT%'); $s.TargetPath = '%TARGET_EXE%'; $s.WorkingDirectory = '%WORKING_DIR%'; $s.WindowStyle = 1; $s.Save()"

:: 7. Create Uninstaller Script
echo.
echo [6/6] Generating Uninstaller...
echo @echo off > uninstall.bat
echo title Uninstall PDF Annotation Studio >> uninstall.bat
echo echo =================================================== >> uninstall.bat
echo echo Are you sure you want to uninstall PDF Annotation Studio? >> uninstall.bat
echo echo This will remove the compiled app, virtual environment, and Start Menu shortcuts. >> uninstall.bat
echo echo =================================================== >> uninstall.bat
echo pause >> uninstall.bat
echo echo Taskkilling app if running... >> uninstall.bat
echo taskkill /F /IM "PDF Annotation Studio.exe" ^>nul 2^>^&1 >> uninstall.bat
echo echo Removing Start Menu shortcuts... >> uninstall.bat
echo rmdir /S /Q "%START_MENU_DIR%" >> uninstall.bat
echo echo Removing compiled builds and cache... >> uninstall.bat
echo rmdir /S /Q "%CD%\build" >> uninstall.bat
echo rmdir /S /Q "%CD%\dist" >> uninstall.bat
echo del /Q "%CD%\PDF Annotation Studio.spec" >> uninstall.bat
echo echo Removing virtual environment... >> uninstall.bat
echo rmdir /S /Q "%CD%\venv" >> uninstall.bat
echo echo Uninstall complete. You may delete this folder. >> uninstall.bat
echo pause >> uninstall.bat
:: Self-delete the uninstaller script after it finishes
echo (goto) 2^>nul ^& del "%%~f0" >> uninstall.bat

:: Create Uninstall Shortcut
set "UNINSTALL_TARGET=%CD%\uninstall.bat"
powershell -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%UNINSTALL_SHORTCUT%'); $s.TargetPath = '%UNINSTALL_TARGET%'; $s.WorkingDirectory = '%CD%'; $s.WindowStyle = 1; $s.Save()"

echo ===================================================
echo Setup Complete!
echo Application compiled to a native Windows binary.
echo Starting application now...
echo ===================================================

:: Start the app directly
start "" "%TARGET_EXE%"
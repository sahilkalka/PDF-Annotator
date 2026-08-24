@echo off
setlocal EnableExtensions

set "APP_NAME=PDF Annotator"
set "APP_DIR=%~dp0"
if "%APP_DIR:~-1%"=="\" set "APP_DIR=%APP_DIR:~0,-1%"

set "ICON_PATH=%APP_DIR%\icon.ico"
set "LAUNCHER=%APP_DIR%\run_pdf_annotator.bat"
set "VENV_PY=%APP_DIR%\.venv\Scripts\python.exe"
set "START_MENU_DIR=%APPDATA%\Microsoft\Windows\Start Menu\Programs\%APP_NAME%"

where py >nul 2>nul
if %errorlevel% equ 0 (
    set "PY_BOOTSTRAP=py -3"
) else (
    where python >nul 2>nul
    if not %errorlevel% equ 0 (
        echo Python 3 not found. Install Python 3 and re-run setup.
        exit /b 1
    )
    set "PY_BOOTSTRAP=python"
)

echo [1/4] Creating virtual environment...
if not exist "%APP_DIR%\.venv\Scripts\python.exe" (
    %PY_BOOTSTRAP% -m venv "%APP_DIR%\.venv"
    if not %errorlevel% equ 0 (
        echo Failed to create virtual environment.
        exit /b 1
    )
) else (
    echo Virtual environment already exists.
)

echo [2/4] Installing dependencies...
"%VENV_PY%" -m pip install --upgrade pip
if not %errorlevel% equ 0 (
    echo Failed to upgrade pip.
    exit /b 1
)

if exist "%APP_DIR%\requirements.txt" (
    "%VENV_PY%" -m pip install -r "%APP_DIR%\requirements.txt"
) else (
    "%VENV_PY%" -m pip install Flask PyMuPDF openpyxl
)
if not %errorlevel% equ 0 (
    echo Dependency installation failed.
    exit /b 1
)

echo [3/4] Creating launcher and shortcuts...
if not exist "%START_MENU_DIR%" mkdir "%START_MENU_DIR%"

powershell -NoProfile -NonInteractive -Command ^
"$ErrorActionPreference='Stop'; ^
$w=New-Object -ComObject WScript.Shell; ^
$appName='PDF Annotator'; ^
$appDir='%APP_DIR%'; ^
$icon='%ICON_PATH%'; ^
$launcher='%LAUNCHER%'; ^
$startMenuDir='%START_MENU_DIR%'; ^
$desktop=[Environment]::GetFolderPath('Desktop'); ^
if(-not (Test-Path $startMenuDir)){ [void](New-Item -ItemType Directory -Path $startMenuDir -Force) }; ^
$runStart=Join-Path $startMenuDir ($appName + '.lnk'); ^
$runDesk=Join-Path $desktop ($appName + '.lnk'); ^
$runTargets=@($runStart,$runDesk); ^
foreach($p in $runTargets){ ^
  $s=$w.CreateShortcut($p); ^
  $s.TargetPath=$launcher; ^
  $s.WorkingDirectory=$appDir; ^
  if(Test-Path $icon){ $s.IconLocation=$icon }; ^
  $s.Description='Run PDF Annotator'; ^
  $s.Save() ^
}; ^
$oldWebStart=Join-Path $startMenuDir ($appName + ' - Open App.lnk'); ^
$oldWebDesk=Join-Path $desktop ($appName + ' - Open App.lnk'); ^
foreach($p in @($oldWebStart,$oldWebDesk)){ if(Test-Path $p){ Remove-Item -Force $p } }"
if not %errorlevel% equ 0 (
    echo Failed to create shortcuts.
    exit /b 1
)

echo [4/4] Setup complete.
echo Start Menu: %START_MENU_DIR%
echo Desktop shortcuts created for %APP_NAME%.
echo.
echo Security notes:
echo - Per-user install only (no admin required)
echo - Localhost binding only (127.0.0.1)
echo - No external startup scripts or registry autoruns

endlocal
exit /b 0

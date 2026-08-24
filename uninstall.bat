@echo off
setlocal EnableExtensions

set "APP_NAME=PDF Annotator"
set "APP_DIR=%~dp0"
if "%APP_DIR:~-1%"=="\" set "APP_DIR=%APP_DIR:~0,-1%"

set "START_MENU_DIR=%APPDATA%\Microsoft\Windows\Start Menu\Programs\%APP_NAME%"
set "DESKTOP_DIR=%USERPROFILE%\Desktop"
set "VENV_DIR=%APP_DIR%\.venv"

echo Uninstalling %APP_NAME% (per-user shortcuts only)...

echo [1/3] Removing shortcuts...
del /f /q "%START_MENU_DIR%\%APP_NAME%.lnk" >nul 2>nul
del /f /q "%DESKTOP_DIR%\%APP_NAME%.lnk" >nul 2>nul

if exist "%START_MENU_DIR%" (
    rmdir "%START_MENU_DIR%" >nul 2>nul
)

echo [2/3] Shortcut cleanup complete.

echo [3/3] Optional environment cleanup...
set "REMOVE_VENV="
set /p REMOVE_VENV=Delete local Python virtual environment (.venv)? [y/N]: 
if /I "%REMOVE_VENV%"=="Y" (
    if exist "%VENV_DIR%" (
        rmdir /s /q "%VENV_DIR%"
        if exist "%VENV_DIR%" (
            echo Failed to remove .venv. Close running Python processes and try again.
            exit /b 1
        )
        echo .venv removed.
    ) else (
        echo .venv not found. Skipping.
    )
) else (
    echo Keeping .venv.
)

echo.
echo Uninstall complete.
echo Removed: Start Menu and Desktop shortcuts for %APP_NAME%.
echo Kept: Application source files in %APP_DIR%.

endlocal
exit /b 0

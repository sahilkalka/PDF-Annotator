@echo off 
title Uninstall PDF Annotation Studio 
echo =================================================== 
echo Are you sure you want to uninstall PDF Annotation Studio? 
echo This will remove the compiled app, virtual environment, and Start Menu shortcuts. 
echo =================================================== 
pause 
echo Taskkilling app if running... 
taskkill /F /IM "PDF Annotation Studio.exe" >nul 2>&1 
echo Removing Start Menu shortcuts... 
rmdir /S /Q "C:\Users\2286369\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\PDF Annotation Studio" 
echo Removing compiled builds and cache... 
rmdir /S /Q "C:\Users\2286369\OneDrive - Cognizant\Documents\AI Projects\PDF Annotator\build" 
rmdir /S /Q "C:\Users\2286369\OneDrive - Cognizant\Documents\AI Projects\PDF Annotator\dist" 
del /Q "C:\Users\2286369\OneDrive - Cognizant\Documents\AI Projects\PDF Annotator\PDF Annotation Studio.spec" 
echo Removing virtual environment... 
rmdir /S /Q "C:\Users\2286369\OneDrive - Cognizant\Documents\AI Projects\PDF Annotator\venv" 
echo Uninstall complete. You may delete this folder. 
pause 
(goto) 2>nul & del "%~f0" 

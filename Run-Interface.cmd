@echo off
setlocal
if not defined CLINIC_RSCRIPT set "CLINIC_RSCRIPT=Rscript.exe"
pushd "%~dp0"
if errorlevel 1 goto drive_error
"%CLINIC_RSCRIPT%" "%~dp0run_app.R"
set "CLINIC_EXIT=%ERRORLEVEL%"
popd
if not "%CLINIC_EXIT%"=="0" echo Interface stopped with an error. Read the message above.
pause
exit /b %CLINIC_EXIT%
:drive_error
echo Cannot open the project folder.
pause
exit /b 1

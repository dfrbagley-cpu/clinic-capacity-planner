@echo off
setlocal
rem Set CLINIC_RSCRIPT to the IT-installed Rscript.exe path, or add R to PATH.
if not defined CLINIC_RSCRIPT set "CLINIC_RSCRIPT=Rscript.exe"
pushd "%~dp0"
if errorlevel 1 goto drive_error
"%CLINIC_RSCRIPT%" "%~dp0run_report.R" --mode basic
set "CLINIC_EXIT=%ERRORLEVEL%"
popd
if not "%CLINIC_EXIT%"=="0" echo Report failed. Read the message above; no successful report is implied.
pause
exit /b %CLINIC_EXIT%
:drive_error
echo Cannot open the shared-drive project folder.
pause
exit /b 1

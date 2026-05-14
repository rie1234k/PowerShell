@echo off

rem ˆø”‚ª‚È‚¢ê‡‚ÍI—¹‚·‚é
if "%~1"=="" (
  echo ˆø”‚ª‚ ‚è‚Ü‚¹‚ñ
  pause
  exit /b
)
pushd %~dp0 
powershell -executionpolicy RemoteSigned -File "CreateLaunchShortcut.ps1" %1


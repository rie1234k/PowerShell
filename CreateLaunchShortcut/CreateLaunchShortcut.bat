@echo off

rem 引数がない場合は終了する
if "%~1"=="" (
  echo 引数がありません
  pause
  exit /b
)
pushd %~dp0 
powershell -executionpolicy RemoteSigned -File "CreateLaunchShortcut.ps1" %1


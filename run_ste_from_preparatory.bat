@echo off
cd /d "%~dp0"
if "%~1"=="" (
  python launch_ste_prep.py
) else (
  python run_ste_from_preparatory.py %*
)
if errorlevel 1 pause

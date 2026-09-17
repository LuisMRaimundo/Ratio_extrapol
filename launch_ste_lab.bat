@echo off
cd /d "%~dp0"
python launch_ste_lab.py
if errorlevel 1 pause

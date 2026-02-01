@echo off
cd /d "%~dp0"
powershell -ExecutionPolicy Bypass -File ".\Setup-Shortcut.ps1"
pause

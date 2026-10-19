@echo off
cd /d "%~dp0"
powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%~dp0Rename-FilesByMkab.ps1" -ConfigPath "%~dp0config.json" -ApplyRename

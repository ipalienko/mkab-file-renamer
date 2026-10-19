@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$e=$null; [void][System.Management.Automation.Language.Parser]::ParseFile('%~dp0Rename-FilesByMkab.ps1',[ref]$null,[ref]$e); if($e.Count){$e | Format-List Message,Extent; exit 1}else{Write-Host 'PowerShell syntax: OK' -ForegroundColor Green; exit 0}"
pause

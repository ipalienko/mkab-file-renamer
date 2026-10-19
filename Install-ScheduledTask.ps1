[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$RunAsUser,

    [string]$TaskName = "MKAB File Renamer",
    [string]$DailyAt = "03:00"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$scriptPath = Join-Path $PSScriptRoot "Rename-FilesByMkab.ps1"
$configPath = Join-Path $PSScriptRoot "config.json"

if (-not (Test-Path -LiteralPath $scriptPath -PathType Leaf)) { throw "Script not found: $scriptPath" }
if (-not (Test-Path -LiteralPath $configPath -PathType Leaf)) { throw "Config not found: $configPath" }

$cred = Get-Credential -UserName $RunAsUser -Message "Enter the password for the Task Scheduler account"
$at = [datetime]::ParseExact($DailyAt, "HH:mm", [System.Globalization.CultureInfo]::InvariantCulture)

$powerShell = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
$arguments = ('-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{0}" -ConfigPath "{1}" -ApplyRename' -f $scriptPath, $configPath)

$action = New-ScheduledTaskAction -Execute $powerShell -Argument $arguments -WorkingDirectory $PSScriptRoot
$trigger = New-ScheduledTaskTrigger -Daily -At $at
$settings = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 2)
$principal = New-ScheduledTaskPrincipal -UserId $RunAsUser -LogonType Password -RunLevel Limited
$task = New-ScheduledTask -Action $action -Trigger $trigger -Settings $settings -Principal $principal

$bstr = [IntPtr]::Zero
try {
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($cred.Password)
    $plainPassword = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    Register-ScheduledTask -TaskName $TaskName -InputObject $task -User $RunAsUser -Password $plainPassword -Force | Out-Null
}
finally {
    if ($bstr -ne [IntPtr]::Zero) {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }
    Remove-Variable plainPassword -ErrorAction SilentlyContinue
}

Write-Host "Task installed: $TaskName" -ForegroundColor Green
Write-Host "Schedule: daily at $DailyAt"
Write-Host "Run as: $RunAsUser"
Write-Host "Remember: the account needs 'Log on as a batch job' and access to the file share." -ForegroundColor Yellow

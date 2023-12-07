$scriptPath = "uploads3.ps1"
$destinationPath = "C:\Windows\System32"
Copy-Item -Path $scriptPath -Destination $destinationPath -Force
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-File $($destinationPath)\$($scriptPath)"
$trigger = New-ScheduledTaskTrigger -AtStartup
Register-ScheduledTask -Action $action -Trigger $trigger -TaskName "ExecutarScriptNoInicio"

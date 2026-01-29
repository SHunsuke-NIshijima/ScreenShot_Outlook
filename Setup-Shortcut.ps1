$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$vbsPath = Join-Path $scriptPath "Run-Hidden.vbs"
$shortcutPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Screenshot-Email.lnk"

$WScriptShell = New-Object -ComObject WScript.Shell
$shortcut = $WScriptShell.CreateShortcut($shortcutPath)

$shortcut.TargetPath = "$env:SystemRoot\System32\wscript.exe"
$shortcut.Arguments = "`"$vbsPath`""
$shortcut.WorkingDirectory = $scriptPath
$shortcut.Hotkey = "CTRL+ALT+S"
$shortcut.Description = "Screenshot to Email"

$shortcut.Save()

Write-Host "Shortcut registered successfully!" -ForegroundColor Green
Write-Host "Hotkey: Ctrl + Alt + S" -ForegroundColor Cyan
Write-Host "Location: $shortcutPath" -ForegroundColor Cyan

$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path

$files = @(
    "Send-Screenshot-Email.ps1",
    "Setup-Shortcut.ps1",
    "config.json"
)

foreach ($file in $files) {
    $filePath = Join-Path $scriptPath $file
    if (Test-Path $filePath) {
        $content = Get-Content $filePath -Raw
        $utf8BOM = New-Object System.Text.UTF8Encoding $true
        [System.IO.File]::WriteAllText($filePath, $content, $utf8BOM)
        Write-Host "Fixed encoding for: $file" -ForegroundColor Green
    }
}

Write-Host "`nAll files have been re-saved with UTF-8 BOM encoding!" -ForegroundColor Cyan

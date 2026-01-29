Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$configPath = Join-Path $scriptPath "config.json"
$config = Get-Content $configPath -Raw -Encoding UTF8 | ConvertFrom-Json

$date = Get-Date -Format "yyyyMMdd"
$imageFolder = Join-Path $scriptPath "image\$date"
if (-not (Test-Path $imageFolder)) {
    New-Item -ItemType Directory -Path $imageFolder -Force | Out-Null
}

$timestamp = Get-Date -Format "HHmmss"
$screenshotPath = Join-Path $imageFolder "screenshot_$timestamp.png"

$screen = [System.Windows.Forms.Screen]::PrimaryScreen
$bounds = $screen.Bounds
$bitmap = New-Object System.Drawing.Bitmap $bounds.Width, $bounds.Height
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$graphics.CopyFromScreen($bounds.Location, [System.Drawing.Point]::Empty, $bounds.Size)
$bitmap.Save($screenshotPath, [System.Drawing.Imaging.ImageFormat]::Png)
$graphics.Dispose()
$bitmap.Dispose()

$outlook = New-Object -ComObject Outlook.Application
$mail = $outlook.CreateItem(0)

$mail.To = $config.mailto
$mail.CC = $config.cc
$mail.Subject = $config.subject
$mail.BodyFormat = 2

$attachment = $mail.Attachments.Add($screenshotPath)
$attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "screenshot")

$mail.HTMLBody = @"
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
</head>
<body>
<p>スクリーンショットを添付します。</p>
<p><img src="cid:screenshot" /></p>
</body>
</html>
"@

$mail.Display()

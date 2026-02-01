Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$configPath = Join-Path $scriptPath "config.json"
$config = Get-Content $configPath -Raw -Encoding UTF8 | ConvertFrom-Json

$date = Get-Date -Format "yyyyMMdd"
$dateFormatted = Get-Date -Format "yyyy/MM/dd"
$imageFolder = Join-Path $scriptPath "image\$date"
if (-not (Test-Path $imageFolder)) {
    New-Item -ItemType Directory -Path $imageFolder -Force | Out-Null
}

$timestamp = Get-Date -Format "HHmmss"
$screenshotPath = Join-Path $imageFolder "screenshot_$timestamp.png"

# 枠の色を取得（デフォルト: Red）
$borderColor = if ($config.PSObject.Properties.Name -contains "selectionBorderColor") {
    $config.selectionBorderColor
} else {
    "Red"
}

# 範囲選択用フォームの作成
$selectionForm = New-Object System.Windows.Forms.Form
$selectionForm.FormBorderStyle = 'None'
$selectionForm.WindowState = 'Maximized'
$selectionForm.TopMost = $true
$selectionForm.Cursor = [System.Windows.Forms.Cursors]::Cross
$selectionForm.BackColor = [System.Drawing.Color]::Black
$selectionForm.Opacity = 0.3

$startPoint = $null
$endPoint = $null
$isDrawing = $false

# マウスダウンイベント
$selectionForm.Add_MouseDown({
    param($sender, $e)
    $script:startPoint = $e.Location
    $script:isDrawing = $true
})

# マウスムーブイベント
$selectionForm.Add_MouseMove({
    param($sender, $e)
    if ($script:isDrawing) {
        $selectionForm.Refresh()
    }
})

# 描画イベント
$selectionForm.Add_Paint({
    param($sender, $e)
    if ($script:isDrawing -and $script:startPoint) {
        $currentPoint = $selectionForm.PointToClient([System.Windows.Forms.Cursor]::Position)
        $x = [Math]::Min($script:startPoint.X, $currentPoint.X)
        $y = [Math]::Min($script:startPoint.Y, $currentPoint.Y)
        $width = [Math]::Abs($currentPoint.X - $script:startPoint.X)
        $height = [Math]::Abs($currentPoint.Y - $script:startPoint.Y)
        
        $rect = New-Object System.Drawing.Rectangle($x, $y, $width, $height)
        $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::$borderColor, 2)
        $e.Graphics.DrawRectangle($pen, $rect)
        $pen.Dispose()
    }
})

# マウスアップイベント
$selectionForm.Add_MouseUp({
    param($sender, $e)
    $script:endPoint = $e.Location
    $script:isDrawing = $false
    $selectionForm.Close()
})

# ESCキーでキャンセル
$selectionForm.Add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq 'Escape') {
        $selectionForm.Close()
        exit
    }
})

[void]$selectionForm.ShowDialog()

# 選択範囲が有効かチェック
if ($null -eq $startPoint -or $null -eq $endPoint) {
    Write-Host "Screenshot cancelled" -ForegroundColor Yellow
    exit
}

# 選択範囲の計算
$x = [Math]::Min($startPoint.X, $endPoint.X)
$y = [Math]::Min($startPoint.Y, $endPoint.Y)
$width = [Math]::Abs($endPoint.X - $startPoint.X)
$height = [Math]::Abs($endPoint.Y - $startPoint.Y)

# 範囲が小さすぎる場合はキャンセル
if ($width -lt 10 -or $height -lt 10) {
    Write-Host "Selected area too small" -ForegroundColor Yellow
    exit
}

# 選択範囲のスクリーンショットを撮影
$bounds = New-Object System.Drawing.Rectangle($x, $y, $width, $height)
$bitmap = New-Object System.Drawing.Bitmap $width, $height
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$graphics.CopyFromScreen($x, $y, 0, 0, $bounds.Size)
$bitmap.Save($screenshotPath, [System.Drawing.Imaging.ImageFormat]::Png)
$graphics.Dispose()
$bitmap.Dispose()

$outlook = New-Object -ComObject Outlook.Application
$mail = $outlook.CreateItem(0)

$mail.To = $config.mailto
$mail.CC = $config.cc

# 件名の設定（変数置換）
$subject = if ($config.PSObject.Properties.Name -contains "subject") {
    $config.subject -replace '\{name\}', $config.name -replace '\{date\}', $dateFormatted
} else {
    "【業務連絡】 （$($config.name)） $dateFormatted"
}
$mail.Subject = $subject
$mail.BodyFormat = 2

$attachment = $mail.Attachments.Add($screenshotPath)
$attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "screenshot")

# メール本文を取得（デフォルト: スクリーンショットを添付します。）
$bodyText = if ($config.PSObject.Properties.Name -contains "body") {
    $config.body
} else {
    "スクリーンショットを添付します。"
}

$mail.HTMLBody = @"
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
</head>
<body>
<p>$bodyText</p>
<p><img src="cid:screenshot" width="1200" /></p>
</body>
</html>
"@

$mail.Display()

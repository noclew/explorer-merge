# Original geometric artwork. Regenerate only when changing the icon.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
$bitmap = New-Object System.Drawing.Bitmap 64,64
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$graphics.Clear([System.Drawing.Color]::Transparent)
$back = New-Object System.Drawing.SolidBrush ([System.Drawing.ColorTranslator]::FromHtml('#D99827'))
$front = New-Object System.Drawing.SolidBrush ([System.Drawing.ColorTranslator]::FromHtml('#FFD166'))
$blue = New-Object System.Drawing.SolidBrush ([System.Drawing.ColorTranslator]::FromHtml('#176BCE'))
$white = New-Object System.Drawing.Pen ([System.Drawing.Color]::White),4
$white.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
$white.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
$graphics.FillRectangle($back,6,8,21,9)
$graphics.FillRectangle($back,6,14,46,32)
$graphics.FillRectangle($front,10,22,47,33)
$graphics.FillEllipse($blue,28,30,32,32)
$graphics.DrawLine($white,36,46,51,46)
$graphics.DrawLine($white,46,40,52,46)
$graphics.DrawLine($white,46,52,52,46)
$handle = $bitmap.GetHicon()
$icon = [System.Drawing.Icon]::FromHandle($handle)
$stream = [IO.File]::Create((Join-Path $PSScriptRoot '..\assets\ExplorerMerge.ico'))
try { $icon.Save($stream) } finally { $stream.Dispose(); $icon.Dispose(); $graphics.Dispose(); $bitmap.Dispose(); $back.Dispose(); $front.Dispose(); $blue.Dispose(); $white.Dispose() }

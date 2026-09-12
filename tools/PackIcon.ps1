param(
    [string]$Source = (Join-Path $PSScriptRoot '..\assets\ExplorerMerge-character.png'),
    [string]$Output = (Join-Path $PSScriptRoot '..\assets\ExplorerMerge-character.ico')
)
# Package the finished artwork as a multi-resolution Windows ICO.
# This performs only format conversion and resampling, not drawing changes.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
$original = [Drawing.Image]::FromFile([IO.Path]::GetFullPath($Source))
$sizes = @(16,24,32,48,64,128,256)
$frames = New-Object 'System.Collections.Generic.List[byte[]]'
try {
    foreach ($size in $sizes) {
        $bitmap = New-Object Drawing.Bitmap $size,$size
        $graphics = [Drawing.Graphics]::FromImage($bitmap)
        $graphics.InterpolationMode = [Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $graphics.PixelOffsetMode = [Drawing.Drawing2D.PixelOffsetMode]::HighQuality
        $graphics.DrawImage($original,0,0,$size,$size)
        $stream = New-Object IO.MemoryStream
        try {
            $bitmap.Save($stream,[Drawing.Imaging.ImageFormat]::Png)
            $frames.Add($stream.ToArray())
        } finally { $stream.Dispose(); $graphics.Dispose(); $bitmap.Dispose() }
    }
    $file = [IO.File]::Create([IO.Path]::GetFullPath($Output))
    $writer = New-Object IO.BinaryWriter $file
    try {
        $writer.Write([uint16]0); $writer.Write([uint16]1); $writer.Write([uint16]$sizes.Count)
        $offset = 6 + 16 * $sizes.Count
        for ($i=0; $i -lt $sizes.Count; $i++) {
            $dimension = if ($sizes[$i] -eq 256) { 0 } else { $sizes[$i] }
            $writer.Write([byte]$dimension); $writer.Write([byte]$dimension)
            $writer.Write([byte]0); $writer.Write([byte]0)
            $writer.Write([uint16]1); $writer.Write([uint16]32)
            $writer.Write([uint32]$frames[$i].Length); $writer.Write([uint32]$offset)
            $offset += $frames[$i].Length
        }
        foreach ($frame in $frames) { $writer.Write($frame) }
    } finally { $writer.Dispose(); $file.Dispose() }
} finally { $original.Dispose() }
Write-Output "Created $Output (16, 24, 32, 48, 64, 128, 256 px)"

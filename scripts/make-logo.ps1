Add-Type -AssemblyName System.Drawing

function Add-RoundedRect {
    param(
        [System.Drawing.Drawing2D.GraphicsPath]$Path,
        [float]$X, [float]$Y, [float]$W, [float]$H, [float]$R
    )
    if ($R -le 0) {
        $Path.AddRectangle([System.Drawing.RectangleF]::new($X, $Y, $W, $H))
        return
    }
    $d = $R * 2
    $Path.AddArc($X,         $Y,         $d, $d, 180, 90)
    $Path.AddArc($X + $W - $d, $Y,         $d, $d, 270, 90)
    $Path.AddArc($X + $W - $d, $Y + $H - $d, $d, $d, 0,   90)
    $Path.AddArc($X,         $Y + $H - $d, $d, $d, 90,  90)
    $Path.CloseFigure()
}

function New-LogoBitmap {
    param([int]$Size)

    [float]$S = $Size

    $bmp = New-Object System.Drawing.Bitmap $Size, $Size
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode       = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.CompositingQuality  = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
    $g.PixelOffsetMode     = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $g.InterpolationMode   = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic

    # ---------- Background: dark rounded square with subtle top sheen ----------
    [float]$corner = $S * 0.22
    $bgPath = [System.Drawing.Drawing2D.GraphicsPath]::new()
    Add-RoundedRect -Path $bgPath -X 0 -Y 0 -W $S -H $S -R $corner

    $bgRect = [System.Drawing.RectangleF]::new(0, 0, $S, $S)
    $bgTop  = [System.Drawing.Color]::FromArgb(255, 24, 24, 27)
    $bgBot  = [System.Drawing.Color]::FromArgb(255, 10, 10, 12)
    $bgBrush = [System.Drawing.Drawing2D.LinearGradientBrush]::new($bgRect, $bgTop, $bgBot, [single]90.0)
    $g.FillPath($bgBrush, $bgPath)

    # Hairline inner border
    [float]$borderThickness = [Math]::Max(1.0, $S * 0.006)
    $borderPen = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(38, 255, 255, 255), $borderThickness)
    $g.DrawPath($borderPen, $bgPath)

    # Soft top highlight
    $sheenRect = [System.Drawing.RectangleF]::new(0, 0, $S, $S * 0.5)
    $sheenPath = [System.Drawing.Drawing2D.GraphicsPath]::new()
    Add-RoundedRect -Path $sheenPath -X 0 -Y 0 -W $S -H ($S * 0.5) -R $corner
    $sheenBrush = [System.Drawing.Drawing2D.LinearGradientBrush]::new(
        $sheenRect,
        [System.Drawing.Color]::FromArgb(20, 255, 255, 255),
        [System.Drawing.Color]::FromArgb(0, 255, 255, 255),
        [single]90.0)
    $g.FillPath($sheenBrush, $sheenPath)

    # ---------- Refinement funnel: three stacked bars, narrowing top to bottom ----------
    [float]$barH   = $S * 0.12
    [float]$gap    = $S * 0.07
    [float]$total  = ($barH * 3) + ($gap * 2)
    [float]$startY = ($S - $total) / 2.0
    [float]$cx     = $S / 2.0
    [float]$barR   = $barH / 2.0   # fully rounded pill ends

    [float]$w1 = $S * 0.62    # dim — raw rows
    [float]$w2 = $S * 0.46    # purple — in progress
    [float]$w3 = $S * 0.28    # gradient — refined output

    # Bar 1 (top, dim gray)
    $b1Path = [System.Drawing.Drawing2D.GraphicsPath]::new()
    Add-RoundedRect -Path $b1Path -X ($cx - $w1 / 2) -Y $startY -W $w1 -H $barH -R $barR
    $b1Brush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 70, 70, 82))
    $g.FillPath($b1Brush, $b1Path)

    # Bar 2 (middle, solid purple)
    [float]$b2Y = $startY + $barH + $gap
    $b2Path = [System.Drawing.Drawing2D.GraphicsPath]::new()
    Add-RoundedRect -Path $b2Path -X ($cx - $w2 / 2) -Y $b2Y -W $w2 -H $barH -R $barR
    $b2Brush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 124, 92, 255))
    $g.FillPath($b2Brush, $b2Path)

    # Bar 3 (bottom, gradient — the refined result)
    [float]$b3Y = $startY + ($barH + $gap) * 2
    [float]$b3X = $cx - $w3 / 2
    $b3Path = [System.Drawing.Drawing2D.GraphicsPath]::new()
    Add-RoundedRect -Path $b3Path -X $b3X -Y $b3Y -W $w3 -H $barH -R $barR
    $b3Rect = [System.Drawing.RectangleF]::new($b3X, $b3Y, $w3, $barH)
    $b3Brush = [System.Drawing.Drawing2D.LinearGradientBrush]::new(
        $b3Rect,
        [System.Drawing.Color]::FromArgb(255, 124, 92, 255),
        [System.Drawing.Color]::FromArgb(255, 236, 72, 152),
        [single]0.0)
    $g.FillPath($b3Brush, $b3Path)

    # Soft glow under the refined bar (only when there's enough pixels)
    if ($Size -ge 64) {
        [float]$glowW = $w3 * 1.35
        [float]$glowH = $barH * 0.55
        [float]$glowX = $cx - $glowW / 2
        [float]$glowY = $b3Y + $barH + $S * 0.012
        $glowPath = [System.Drawing.Drawing2D.GraphicsPath]::new()
        Add-RoundedRect -Path $glowPath -X $glowX -Y $glowY -W $glowW -H $glowH -R ($glowH / 2)
        $glowRect = [System.Drawing.RectangleF]::new($glowX, $glowY, $glowW, $glowH)
        $glowBrush = [System.Drawing.Drawing2D.LinearGradientBrush]::new(
            $glowRect,
            [System.Drawing.Color]::FromArgb(75, 124, 92, 255),
            [System.Drawing.Color]::FromArgb(75, 236, 72, 152),
            [single]0.0)
        $g.FillPath($glowBrush, $glowPath)
        $glowBrush.Dispose()
        $glowPath.Dispose()
    }

    $g.Dispose()
    $bgBrush.Dispose()
    $borderPen.Dispose()
    $sheenBrush.Dispose()
    $b1Brush.Dispose()
    $b2Brush.Dispose()
    $b3Brush.Dispose()
    $bgPath.Dispose()
    $sheenPath.Dispose()
    $b1Path.Dispose()
    $b2Path.Dispose()
    $b3Path.Dispose()

    return $bmp
}

function Get-PngBytes {
    param([System.Drawing.Bitmap]$Bitmap)
    $ms = [System.IO.MemoryStream]::new()
    $Bitmap.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    return $ms.ToArray()
}

function Write-Ico {
    param([hashtable]$Entries, [string]$OutPath)
    $keys = $Entries.Keys | Sort-Object
    $count = $keys.Count

    $fs = [System.IO.File]::Open($OutPath, [System.IO.FileMode]::Create)
    $bw = [System.IO.BinaryWriter]::new($fs)

    $bw.Write([UInt16]0)
    $bw.Write([UInt16]1)
    $bw.Write([UInt16]$count)

    $offset = 6 + ($count * 16)

    foreach ($size in $keys) {
        $data = $Entries[$size]
        $w = if ($size -ge 256) { 0 } else { [Byte]$size }
        $h = $w
        $bw.Write([Byte]$w)
        $bw.Write([Byte]$h)
        $bw.Write([Byte]0)
        $bw.Write([Byte]0)
        $bw.Write([UInt16]1)
        $bw.Write([UInt16]32)
        $bw.Write([UInt32]$data.Length)
        $bw.Write([UInt32]$offset)
        $offset += $data.Length
    }

    foreach ($size in $keys) {
        $data = [byte[]]$Entries[$size]
        $fs.Write($data, 0, $data.Length)
    }

    $bw.Flush()
    $bw.Close()
    $fs.Close()
}

$assetsDir = Join-Path $PSScriptRoot "..\Assets"
if (-not (Test-Path $assetsDir)) {
    New-Item -ItemType Directory -Path $assetsDir | Out-Null
}

$sizes = 16, 24, 32, 48, 64, 128, 256
$entries = @{}

foreach ($size in $sizes) {
    $bmp = New-LogoBitmap -Size $size
    $entries[$size] = Get-PngBytes -Bitmap $bmp
    if ($size -eq 256 -or $size -eq 64) {
        $pngPath = Join-Path $assetsDir "logo-$size.png"
        $bmp.Save($pngPath, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    $bmp.Dispose()
}

$icoPath = Join-Path $assetsDir "DataRefineX.ico"
Write-Ico -Entries $entries -OutPath $icoPath
Write-Host "Wrote $icoPath ($([int]((Get-Item $icoPath).Length / 1KB)) KB, $($sizes.Count) sizes)"

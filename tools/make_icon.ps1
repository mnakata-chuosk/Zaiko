param([string]$OutIco, [string]$Preview)
# 在庫表アイコン：段ボール箱のタイルに「Za」。各サイズを個別に描画し、PNG 形式のエントリで .ico を作る
Add-Type -AssemblyName System.Drawing

function C([int]$r, [int]$g, [int]$b, [int]$a = 255) { [Drawing.Color]::FromArgb($a, $r, $g, $b) }

function Draw-Icon([int]$size) {
    $bmp = New-Object Drawing.Bitmap($size, $size, [Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'; $g.TextRenderingHint = 'AntiAliasGridFit'; $g.PixelOffsetMode = 'HighQuality'
    $g.Clear([Drawing.Color]::Transparent)

    $small = $size -le 24
    $depth = if ($small) { 0 } else { [Math]::Max(2, [int][Math]::Round($size * 0.12)) }
    $w = $size - $depth
    $front = New-Object Drawing.Rectangle(0, $depth, $w, ($size - $depth))

    if ($depth -gt 0) {
        # 上面（フラップ）と右側面
        $top = [Drawing.PointF[]]@((New-Object Drawing.PointF(0, $depth)), (New-Object Drawing.PointF($depth, 0)), (New-Object Drawing.PointF($size, 0)), (New-Object Drawing.PointF($w, $depth)))
        $g.FillPolygon((New-Object Drawing.SolidBrush((C 232 200 150))), $top)
        $side = [Drawing.PointF[]]@((New-Object Drawing.PointF($w, $depth)), (New-Object Drawing.PointF($size, 0)), (New-Object Drawing.PointF($size, ($size - $depth))), (New-Object Drawing.PointF($w, $size)))
        $g.FillPolygon((New-Object Drawing.SolidBrush((C 128 86 40))), $side)
    }

    # 正面（段ボールのグラデーション）
    $br = New-Object Drawing.Drawing2D.LinearGradientBrush($front, (C 218 174 114), (C 168 118 60), 90.0)
    $g.FillRectangle($br, $front)

    # 梱包テープ（上面から正面上部へ垂れる帯）
    $tapeW = [Math]::Max(2, [int][Math]::Round($w * 0.22))
    $tapeX = [int](($w - $tapeW) / 2)
    $tapeH = [int][Math]::Round(($size - $depth) * 0.22)
    $tape = C 196 152 92
    if ($depth -gt 0) {
        $tTop = [Drawing.PointF[]]@((New-Object Drawing.PointF($tapeX, $depth)), (New-Object Drawing.PointF(($tapeX + $depth), 0)), (New-Object Drawing.PointF(($tapeX + $tapeW + $depth), 0)), (New-Object Drawing.PointF(($tapeX + $tapeW), $depth)))
        $g.FillPolygon((New-Object Drawing.SolidBrush((C 205 165 108))), $tTop)
    }
    $g.FillRectangle((New-Object Drawing.SolidBrush($tape)), $tapeX, $depth, $tapeW, $tapeH)

    # 縁取り
    if ($size -ge 32) {
        $pen = New-Object Drawing.Pen((C 110 72 30 120), [single][Math]::Max(1, $size / 64))
        $g.DrawRectangle($pen, 0, $depth, ($w - 1), ($size - $depth - 1))
    }

    # 文字
    $fontSize = if ($small) { $size * 0.74 } else { $size * 0.40 }
    $style = if ($small) { [Drawing.FontStyle]::Bold } else { [Drawing.FontStyle]::Regular }
    $font = New-Object Drawing.Font('Segoe UI Semibold', [single]$fontSize, $style, [Drawing.GraphicsUnit]::Pixel)
    $sf = New-Object Drawing.StringFormat; $sf.Alignment = 'Center'; $sf.LineAlignment = 'Center'
    $textRect = New-Object Drawing.RectangleF(0, ($depth + $tapeH * 0.45), $w, ($size - $depth - $tapeH * 0.45))
    if ($small) { $textRect = New-Object Drawing.RectangleF(-2, 1, ($w + 4), $size) }
    $g.DrawString('Za', $font, (New-Object Drawing.SolidBrush((C 58 34 10))), $textRect, $sf)

    $g.Dispose()
    return $bmp
}

$sizes = 16, 24, 32, 48, 64, 128, 256
$pngs = @()
foreach ($s in $sizes) {
    $bmp = Draw-Icon $s
    $ms = New-Object IO.MemoryStream
    $bmp.Save($ms, [Drawing.Imaging.ImageFormat]::Png)
    $pngs += , @($s, $ms.ToArray())
    $bmp.Dispose()
}

# ICO（ICONDIR + ICONDIRENTRY×n + PNG データ）
$fs = [IO.File]::Create($OutIco)
$bw = New-Object IO.BinaryWriter($fs)
$bw.Write([uint16]0); $bw.Write([uint16]1); $bw.Write([uint16]$pngs.Count)
$offset = 6 + 16 * $pngs.Count
foreach ($p in $pngs) {
    $s = $p[0]; $data = $p[1]
    $bw.Write([byte]($(if ($s -ge 256) { 0 } else { $s }))); $bw.Write([byte]($(if ($s -ge 256) { 0 } else { $s })))
    $bw.Write([byte]0); $bw.Write([byte]0)
    $bw.Write([uint16]1); $bw.Write([uint16]32)
    $bw.Write([uint32]$data.Length); $bw.Write([uint32]$offset)
    $offset += $data.Length
}
foreach ($p in $pngs) { $bw.Write([byte[]]$p[1]) }
$bw.Close()

# プレビュー（各サイズを実寸で並べる＋拡大）
$pv = New-Object Drawing.Bitmap(560, 290); $g = [Drawing.Graphics]::FromImage($pv); $g.Clear([Drawing.Color]::White)
$x = 10
foreach ($s in 16, 24, 32, 48, 64, 128) {
    $img = Draw-Icon $s; $g.DrawImage($img, $x, 10, $s, $s); $g.DrawString("$s", (New-Object Drawing.Font('Yu Gothic UI', 8)), [Drawing.Brushes]::Gray, $x, (14 + $s)); $x += $s + 18
}
$img = Draw-Icon 256; $g.DrawImage($img, 10, 160 - 130 + 120, 128, 128)
$g.DrawString('256（縮小表示）', (New-Object Drawing.Font('Yu Gothic UI', 8)), [Drawing.Brushes]::Gray, 145, 260)
# 既存アイコンと並べた見え方
$x = 300
foreach ($f in 'C:\Source\Repos\受発注管理\受発注管理\Ju.ico', 'C:\Source\Repos\取入表変換\取入表変換\TX.ico') { $ico = New-Object Drawing.Icon($f, 64, 64); $g.DrawIcon($ico, (New-Object Drawing.Rectangle($x, 175, 64, 64))); $x += 80 }
$g.DrawImage((Draw-Icon 64), $x, 175, 64, 64)
$pv.Save($Preview)
"ok: $OutIco ($((Get-Item $OutIco).Length) bytes)"

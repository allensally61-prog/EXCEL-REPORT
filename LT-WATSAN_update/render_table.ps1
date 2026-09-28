# Draws a pipeline-summary table (rows from a JSON file) into a vector EMF.
# Called by update_schematic.py:  powershell -File render_table.ps1 <rows.json> <out.emf>
param([string]$JsonPath, [string]$OutPath, [double]$Aspect = 0)   # Aspect = height/width of the box on the drawing
Add-Type -AssemblyName System.Drawing

$rows = Get-Content -Raw -Encoding UTF8 $JsonPath | ConvertFrom-Json
$cols = @(40, 380, 208)                     # S/N | DESCRIPTION | PIPELINE LENGTH (m)
$W = ($cols | Measure-Object -Sum).Sum
$RH = 13
if ($Aspect -gt 0) { $RH = [math]::Round($W * $Aspect / $rows.Count, 2) }   # rows fill the old table box
$FS = [math]::Round($RH * 0.75, 1)   # font follows row height; shrunk below if the longest description would not fit
$H = $rows.Count * $RH

$ref = New-Object System.Drawing.Bitmap 1, 1
$g0 = [System.Drawing.Graphics]::FromImage($ref)
$hdc = $g0.GetHdc()
# no frame given: the EMF bounds are taken from what is drawn, so there is no DPI-dependent padding
$mf = New-Object System.Drawing.Imaging.Metafile($OutPath, $hdc, [System.Drawing.Imaging.EmfType]::EmfOnly)
$g0.ReleaseHdc($hdc); $g0.Dispose(); $ref.Dispose()
$g = [System.Drawing.Graphics]::FromImage($mf)
$g.PageUnit = [System.Drawing.GraphicsUnit]::Point

function Brush($hex) { New-Object System.Drawing.SolidBrush ([System.Drawing.ColorTranslator]::FromHtml($hex)) }
$pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::Black), 0.5
$longest = ($rows | Where-Object { $_.kind -in 'section', 'pipe' } | ForEach-Object { [string]$_.text }) + 'PIPELINE LENGTH (m)'
do {
    $fB = New-Object System.Drawing.Font 'Calibri', $FS, ([System.Drawing.FontStyle]::Bold), ([System.Drawing.GraphicsUnit]::Point)
    $widest = ($longest | ForEach-Object { $g.MeasureString($_, $fB).Width } | Measure-Object -Maximum).Maximum
    if ($widest -gt $cols[1] - 6) { $FS -= 0.25 }
} while ($widest -gt $cols[1] - 6 -and $FS -gt 5)
$fN = New-Object System.Drawing.Font 'Calibri', $FS, ([System.Drawing.FontStyle]::Regular), ([System.Drawing.GraphicsUnit]::Point)
$fB = New-Object System.Drawing.Font 'Calibri', $FS, ([System.Drawing.FontStyle]::Bold), ([System.Drawing.GraphicsUnit]::Point)
$black = Brush '#000000'; $green = Brush '#00B050'; $red = Brush '#FF0000'
$fill = @{ title = (Brush '#DDEBF7'); head = (Brush '#BDD7EE'); grand = (Brush '#9BC2E6') }
$L = New-Object System.Drawing.StringFormat; $L.LineAlignment = 'Center'; $L.Trimming = 'EllipsisCharacter'
$C = New-Object System.Drawing.StringFormat; $C.Alignment = 'Center'; $C.LineAlignment = 'Center'

function Cell($x, $y, $w, $text, $font, $brush, $fmt, $bg) {
    $r = New-Object System.Drawing.RectangleF $x, $y, $w, $RH
    if ($bg) { $g.FillRectangle($bg, $r) }
    $g.DrawRectangle($pen, $x, $y, $w, $RH)
    if ($text) {
        $t = New-Object System.Drawing.RectangleF ($x + 2), $y, ($w - 4), $RH
        $g.DrawString([string]$text, $font, $brush, $t, $fmt)
    }
}

$g.FillRectangle([System.Drawing.Brushes]::White, 0, 0, $W, $H)
$y = 0
$x1 = $cols[0]; $x2 = $cols[0] + $cols[1]
foreach ($r in $rows) {
    switch ($r.kind) {
        'title'   { Cell 0 $y $W $r.text $fB $black $C $fill.title }
        'head'    { Cell 0 $y $cols[0] 'S/N' $fB $black $C $fill.head
                    Cell $x1 $y $cols[1] 'DESCRIPTION' $fB $black $C $fill.head
                    Cell $x2 $y $cols[2] 'PIPELINE LENGTH (m)' $fB $black $C $fill.head }
        'section' { Cell 0 $y $cols[0] $r.sn $fB $black $C $null
                    Cell $x1 $y $cols[1] $r.text $fB $black $L $null
                    Cell $x2 $y $cols[2] '' $fN $black $C $null }
        'pipe'    { Cell 0 $y $cols[0] '' $fN $black $C $null
                    Cell $x1 $y $cols[1] $r.text $fN $black $L $null
                    Cell $x2 $y $cols[2] $r.len $fN $black $C $null }
        'sub'     { Cell 0 $y $cols[0] '' $fN $black $C $null
                    Cell $x1 $y $cols[1] 'Sub-total' $fB $green $C $null
                    Cell $x2 $y $cols[2] $r.len $fB $green $C $null }
        'grand'   { Cell 0 $y $cols[0] '' $fB $red $C $fill.grand
                    Cell $x1 $y $cols[1] $r.text $fB $red $C $fill.grand
                    Cell $x2 $y $cols[2] $r.len $fB $red $C $fill.grand }
        default   { Cell 0 $y $cols[0] '' $fN $black $C $null
                    Cell $x1 $y $cols[1] '' $fN $black $C $null
                    Cell $x2 $y $cols[2] '' $fN $black $C $null }
    }
    $y += $RH
}
$g.Dispose(); $mf.Dispose()
Write-Output "font $FS pt"
Write-Output "$W $H"

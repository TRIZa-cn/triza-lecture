$ErrorActionPreference = "Stop"
$D = "D:\doc\ZGCA\Lectures\TRIZ\PPT"
$src = Join-Path $D "TRIZ_W2A_矛盾解剖_15min_优化版.pptx"
$dst = Join-Path $D "TRIZ_W2A_矛盾解剖_15min_优化版_自动解说.pptx"
$aud = Join-Path $D "narration_opt"
$meta = Get-Content (Join-Path $D "_narration_opt.json") -Raw -Encoding UTF8 | ConvertFrom-Json

Copy-Item -LiteralPath $src -Destination $dst -Force
Write-Output "copied"

$app = New-Object -ComObject PowerPoint.Application
$ok = $false
for ($i = 0; $i -lt 60; $i++) {
    try { $null = $app.Version; $ok = $true; break } catch { Start-Sleep -Seconds 5 }
}
if (-not $ok) { Write-Output "COM timeout"; exit 2 }
Write-Output "COM ready"

function Retry([scriptblock]$sb) {
    for ($i = 0; $i -lt 40; $i++) {
        try { return & $sb } catch { Start-Sleep -Milliseconds 700 }
    }
    throw "retry exhausted"
}

$pres = Retry { $app.Presentations.Open($dst, 0, 0, -1) }
$n = 0
foreach ($row in $meta.rows) {
    $slide = $pres.Slides.Item([int]$row.slide)
    $file = Join-Path $aud $row.file
    $shape = Retry { $slide.Shapes.AddMediaObject2($file, 0, -1, 975, 562, 18, 18) }
    $shape.Left = 975; $shape.Top = 562; $shape.Width = 18; $shape.Height = 18
    $p = $shape.AnimationSettings
    $p.PlaySettings.PlayOnEntry = -1
    $p.PlaySettings.HideWhileNotPlaying = -1
    $p.PlaySettings.LoopUntilStopped = 0
    $t = $slide.SlideShowTransition
    $t.AdvanceOnTime = -1
    $t.AdvanceTime = [double]$row.dur + 0.6
    $n++
}
Retry { $pres.Save() } | Out-Null
$pres.Close()
$app.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
Write-Output "ASM placed=$n slides=$($meta.rows.Count)"
Write-Output "done"
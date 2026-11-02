$ErrorActionPreference = "Stop"
$root = "D:\doc\ZGCA\Lectures\TRIZ\PPT"
$info = Get-Content (Join-Path $root "_narration.json") -Raw -Encoding UTF8 | ConvertFrom-Json

function Wait-Com {
  for ($i = 0; $i -lt 60; $i++) {
    try { $a = New-Object -ComObject PowerPoint.Application; $null = $a.Version; return $a }
    catch { Start-Sleep -Seconds 5 }
  }
  throw "COM timeout: PowerPoint stayed busy"
}
function Retry([scriptblock]$sb) {
  for ($i = 0; $i -lt 12; $i++) {
    try { return (& $sb) } catch {
      $m = [string]$_.Exception.Message
      if ($m -match "rejected by callee|0x80010001|Call was rejected") { Start-Sleep -Milliseconds 900; continue }
      throw
    }
  }
  throw "COM retry exhausted"
}

"waiting for COM ..."
$app = Wait-Com
"COM ready"
try {
  foreach ($tag in @("A", "B")) {
    $srcName = if ($tag -eq "A") { "TRIZ_W2A_矛盾解剖_15min.pptx" } else { "TRIZ_W2B_盒子与矛盾论_15min.pptx" }
    $dstName = [IO.Path]::GetFileNameWithoutExtension($srcName) + "_自动解说.pptx"
    $src = Join-Path $root $srcName
    $dst = Join-Path $root $dstName
    Copy-Item -LiteralPath $src -Destination $dst -Force
    $prs = Retry { $app.Presentations.Open($dst, 0, 0, -1) }
    $n = 0
    foreach ($r in $info.$tag.rows) {
      if (-not $r.file) { continue }
      $sl = $prs.Slides.Item([int]$r.slide)
      $shp = Retry { $sl.Shapes.AddMediaObject2([string]$r.file, 0, -1, 980, 545, 20, 20) }
      $shp.AnimationSettings.PlaySettings.PlayOnEntry = -1
      $shp.AnimationSettings.PlaySettings.HideWhileNotPlaying = -1
      $shp.AnimationSettings.PlaySettings.LoopUntilStopped = 0
      $t = $sl.SlideShowTransition
      $t.AdvanceOnTime = -1
      $t.AdvanceTime = [single]([double]$r.dur + 0.6)
      $n++
    }
    $prs.Save()
    "ASM $tag placed=$n slides=$($prs.Slides.Count) -> $dstName"
    $prs.Close()
  }
} finally {
  try { $app.Quit() } catch {}
}
"done"
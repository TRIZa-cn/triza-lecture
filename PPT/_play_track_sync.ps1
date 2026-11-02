$ErrorActionPreference = "Continue"
$base = "D:\doc\ZGCA\Lectures\TRIZ\PPT"
$log  = "$base\_sync_play.log"
function L($m) { Add-Content -Path $log -Value ("[{0}] {1}" -f (Get-Date -Format "HH:mm:ss"), $m) -Encoding utf8 }

$cueData = Get-Content -Raw -Encoding UTF8 "$base\_cue_v2.json" | ConvertFrom-Json
$cue   = @($cueData.cues)
$total = [double]$cueData.total
$pptx  = "$base\TRIZa_AI4Science_15min_v2.pptx"
$wav   = "$base\TRIZa_AI4Science_15min_v2_Andrew_8pct.wav"

Set-Content -Path $log -Value ("start sync playback  cues=" + $cue.Count + " total=" + $total) -Encoding utf8
L ("wav=" + $wav + "  exists=" + (Test-Path $wav))

# ---- player -------------------------------------------------------------
$player = $null
try {
  $player = New-Object -ComObject WMPlayer.OCX
  $player.settings.autoStart = $false
  $player.settings.volume = 100
  $player.URL = $wav
  L "player=WMPlayer (position-accurate)"
} catch {
  $player = $null
  L ("player=WMPlayer failed: " + $_.Exception.Message + " -> SoundPlayer+stopwatch")
}

# ---- slideshow ----------------------------------------------------------
$pp = $null; $pres = $null; $show = $false
try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pp.Visible = -1
  $pres = $pp.Presentations.Open($pptx, -1, 0, -1)
  $pres.SlideShowSettings.Run()
  $show = $true
  L "slideshow started"
} catch { L ("slideshow failed: " + $_.Exception.Message) }

Start-Sleep -Milliseconds 900
$shown = 0
try { $pres.SlideShowWindow.View.GotoSlide(1); $shown = 1; L "slide 1" } catch { L "goto slide1 failed" }

# ---- start audio --------------------------------------------------------
$sw = [Diagnostics.Stopwatch]::StartNew()
$sp = $null
if ($player) {
  $player.controls.play()
} else {
  $sp = New-Object System.Media.SoundPlayer $wav
  $sp.Play()
}
L "audio started"

# ---- sync loop ----------------------------------------------------------
$guard = 0
while ($guard -lt 4200) {
  $guard++
  if ($player) {
    try { $pos = [double]$player.currentPosition } catch { $pos = $sw.Elapsed.TotalSeconds }
    $state = 0
    try { $state = [int]$player.playState } catch {}
    if ($pos -le 0) { $pos = $sw.Elapsed.TotalSeconds }
  } else {
    $pos = $sw.Elapsed.TotalSeconds
    $state = 3
  }

  while ($shown -lt $cue.Count -and $pos -ge [double]$cue[$shown].start) {
    $n = $cue[$shown].slide
    try {
      if ($show) { $pres.SlideShowWindow.View.GotoSlide([int]$n) }
      $shown = $n
      L ("slide " + $n + "  at " + [math]::Round($pos,1) + "s  (cue " + $cue[$shown-1].start + ")")
    } catch {
      L ("goto " + $n + " failed: " + $_.Exception.Message)
      try { $pres.SlideShowSettings.Run(); $pres.SlideShowWindow.View.GotoSlide([int]$n); $shown = $n } catch { L "restart failed" }
    }
  }

  if ($pos -ge $total) { L ("reached end at " + [math]::Round($pos,1) + "s"); break }
  if ($state -eq 1 -and $pos -gt 5) { L ("player stopped at " + [math]::Round($pos,1) + "s"); break }
  Start-Sleep -Milliseconds 200
}

L ("done. slides_shown=" + $shown + " elapsed=" + [math]::Round($sw.Elapsed.TotalSeconds,1))
"--- CLEANUP ---"
try { if ($show) { $pres.SlideShowWindow.View.Exit() } } catch {}
try { if ($pres) { $pres.Close() } } catch {}
try { if ($pp) { $pp.Quit() } } catch {}
try { if ($player) { $player.controls.stop(); $player.close() } } catch {}
try { if ($sp) { $sp.Stop(); $sp.Dispose() } } catch {}
L "closed"
"--- CLEANUP DONE ---"

$ErrorActionPreference = "Continue"
$base = "D:\doc\ZGCA\Lectures\TRIZ\PPT"
$log = "$base\_edge_v2_play.log"
function L($m) { Add-Content -Path $log -Value ("[{0}] {1}" -f (Get-Date -Format "HH:mm:ss"), $m) -Encoding utf8 }
Set-Content -Path $log -Value "start (en-US-AndrewNeural -8%, v2 deck, TRIZ=treez)" -Encoding utf8
$d = Get-Content -Raw -Encoding UTF8 "$base\_notes_tts_v2.json" | ConvertFrom-Json
$file = [string]$d.file
$files = @(Get-ChildItem "$base\_edge_v2" -Filter "s*.wav" | Sort-Object Name)
L ("slides=" + $d.slides + " wavs=" + $files.Count + " file=" + $file)
if ($files.Count -lt 2) { L "ABORT: wav missing"; exit 1 }
if (-not (Test-Path $file)) { L "ABORT: pptx missing"; exit 1 }
$pp = $null; $pres = $null; $show = $false
try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pp.Visible = -1
  $pres = $pp.Presentations.Open($file, -1, 0, -1)
  $pres.SlideShowSettings.Run()
  $show = $true
  L "slideshow started"
} catch { L ("slideshow failed: " + $_.Exception.Message) }
for ($i = 0; $i -lt $files.Count; $i++) {
  if ($show) {
    try { if ($i -gt 0) { $pres.SlideShowWindow.View.Next() } }
    catch {
      L ("next failed at " + ($i+1) + " -> restart")
      try { $pres.SlideShowSettings.Run(); $pres.SlideShowWindow.View.GotoSlide($i+1); L "show restarted" } catch { L ("restart failed: " + $_.Exception.Message) }
    }
  }
  Start-Sleep -Milliseconds 600
  L ("play " + ($i+1) + "/" + $files.Count + " " + $files[$i].Name)
  try {
    $pl = New-Object System.Media.SoundPlayer $files[$i].FullName
    $pl.PlaySync()
    $pl.Dispose()
  } catch { L ("play failed: " + $_.Exception.Message) }
  Start-Sleep -Milliseconds 600
}
L "done"
try { if ($show) { $pres.SlideShowWindow.View.Exit() } } catch {}
try { if ($pres) { $pres.Close() } } catch {}
try { if ($pp) { $pp.Quit() } } catch {}
L "closed"

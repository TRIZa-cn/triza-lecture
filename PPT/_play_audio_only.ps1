$ErrorActionPreference = "Continue"
$base = "D:\doc\ZGCA\Lectures\TRIZ\PPT"
$log  = "$base\_audio_only.log"
$wav  = "$base\TRIZa_AI4Science_15min_v2_Andrew_8pct.wav"
function L($m) { Add-Content -Path $log -Value ("[{0}] {1}" -f (Get-Date -Format "HH:mm:ss"), $m) -Encoding utf8 }
Set-Content -Path $log -Value ("audio-only playback  " + $wav) -Encoding utf8
if (-not (Test-Path $wav)) { L "ABORT: wav missing"; exit 1 }
$sw = [Diagnostics.Stopwatch]::StartNew()
L "play start"
try {
  $p = New-Object System.Media.SoundPlayer $wav
  $p.PlaySync()
  $p.Dispose()
  L ("play done  elapsed=" + [math]::Round($sw.Elapsed.TotalSeconds,1) + "s")
} catch {
  L ("play failed: " + $_.Exception.Message)
}
L "exit"

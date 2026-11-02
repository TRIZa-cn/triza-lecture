$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Speech
$base = "D:\doc\ZGCA\Lectures\TRIZ"
$d = Get-Content -Raw -Encoding UTF8 "$base\_notes_dump.json" | ConvertFrom-Json
$notes = @($d.notes)
$sp = New-Object System.Speech.Synthesis.SpeechSynthesizer
$sp.SelectVoice("Microsoft Zira Desktop")
$sp.Rate = -1
for ($i = 0; $i -lt $notes.Count; $i++) {
  $sp.SetOutputToWaveFile((Join-Path "$base\_raw" ("s{0:d2}.wav" -f ($i+1))))
  $sp.Speak([string]$notes[$i])
}
$sp.SetOutputToNull(); $sp.Dispose()
"resynth done"

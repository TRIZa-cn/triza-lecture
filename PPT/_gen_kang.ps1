$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Media.SpeechSynthesis.SpeechSynthesizer, Windows.Media, ContentType=WindowsRuntime]
$null = [Windows.Storage.Streams.DataReader, Windows.Storage.Streams, ContentType=WindowsRuntime]
$asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
  $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
})[0]
function Await($op, $resultType) {
  $t = $asTaskGeneric.MakeGenericMethod($resultType).Invoke($null, @($op))
  $t.Wait(-1) | Out-Null
  $t.Result
}
$base = "D:\doc\ZGCA\Lectures\TRIZ"
$out = Join-Path $base "_kang"
New-Item -ItemType Directory -Force -Path $out | Out-Null
$d = Get-Content -Raw -Encoding UTF8 "$base\_notes_dump.json" | ConvertFrom-Json
$notes = @($d.notes)
$synth = New-Object Windows.Media.SpeechSynthesis.SpeechSynthesizer
$synth.Voice = ([Windows.Media.SpeechSynthesis.SpeechSynthesizer]::AllVoices | Where-Object { $_.DisplayName -like "*Kangkang*" } | Select-Object -First 1)
$synth.Options.SpeakingRate = 1.15
"voice=" + $synth.Voice.DisplayName + " rate=" + $synth.Options.SpeakingRate
for ($i = 0; $i -lt $notes.Count; $i++) {
  $text = [string]$notes[$i]
  $p = Join-Path $out ("s{0:d2}.wav" -f ($i+1))
  $stream = Await ($synth.SynthesizeTextToStreamAsync($text)) ([Windows.Media.SpeechSynthesis.SpeechSynthesisStream])
  $ns = [System.IO.WindowsRuntimeStreamExtensions]::AsStreamForRead($stream)
  $fs = [System.IO.File]::Create($p)
  $ns.CopyTo($fs); $fs.Close(); $ns.Close()
  "  s{0:d2} {1} bytes" -f ($i+1), (Get-Item $p).Length
}
"generated=" + $notes.Count

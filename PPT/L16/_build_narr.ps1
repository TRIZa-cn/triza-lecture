$ErrorActionPreference = "Stop"
$src  = "D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZ_W1_矛盾15m.pptx"
$dst  = "D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZ_W1_矛盾15m_演播版.pptx"
$mp3  = "D:\doc\ZGCA\Lectures\TRIZ\PPT\LWY录音20261030\_嵌入MP3"
$json = "D:\doc\ZGCA\Lectures\TRIZ\PPT\L16\_durations.json"
Copy-Item $src $dst -Force
$dur = Get-Content $json -Raw -Encoding UTF8 | ConvertFrom-Json

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.DisplayAlerts = 1
$pres = $ppt.Presentations.Open($dst, 0, 0, -1)
$W = $pres.PageSetup.SlideWidth
$H = $pres.PageSetup.SlideHeight
Write-Output ("slides={0}  canvas={1:N0}x{2:N0}pt" -f $pres.Slides.Count, $W, $H)

$total = 0
for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    $name = ("p{0:D2}.mp3" -f $i)
    $file = Join-Path $mp3 $name
    if (-not (Test-Path $file)) { Write-Output ("missing " + $name); continue }
    $d = $dur.PSObject.Properties[$name].Value
    $sl = $pres.Slides.Item($i)
    $shp = $sl.Shapes.AddMediaObject2($file, 0, -1, $W + 60, $H + 60)
    $shp.AnimationSettings.PlaySettings.PlayOnEntry = -1
    $shp.AnimationSettings.PlaySettings.HideWhileNotPlaying = -1
    $shp.AnimationSettings.PlaySettings.LoopUntilStopped = 0
    $sl.SlideShowTransition.AdvanceOnTime = -1
    $sl.SlideShowTransition.AdvanceTime = [single]($d + 1.0)
    $total += $d
    Write-Output ("p{0,2} {1} {2,6:N2}s advance={3,6:N2}s shape={4}" -f $i, $name, $d, ($d + 1.0), $shp.Name)
}
$pres.SlideShowSettings.ShowWithAnimation = -1
$pres.SlideShowSettings.ShowWithNarration = -1
$pres.Save()
$pages = $pres.Slides.Count
$pres.Close()
$ppt.Quit()
Start-Sleep -Milliseconds 800
Write-Output ("TOTAL narration {0:N1}s = {1:N2} min ; with advance padding ~{2:N2} min" -f $total, ($total/60), (($total + $pages)/60))
Write-Output ("saved size = {0:N2} MB" -f ((Get-Item $dst).Length/1MB))
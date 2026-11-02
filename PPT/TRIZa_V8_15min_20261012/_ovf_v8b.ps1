param([string]$path)
$app = New-Object -ComObject PowerPoint.Application
$pres = $app.Presentations.Open($path, $true, $false, $false)
$n = 0
foreach ($sl in $pres.Slides) {
  foreach ($sh in $sl.Shapes) {
    if ($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1) {
      $tr = $sh.TextFrame2.TextRange
      $bh = $tr.BoundHeight; $bw = $tr.BoundWidth
      if ($bh -gt ($sh.Height + 4) -or $bw -gt ($sh.Width + 4)) {
        $n++
        $t = $tr.Text -replace "[\r\n]+", " "
        if ($t.Length -gt 50) { $t = $t.Substring(0,50) }
        Write-Output ("S" + $sl.SlideIndex + " H=" + [math]::Round($bh,1) + "/" + [math]::Round($sh.Height,1) + " W=" + [math]::Round($bw,1) + "/" + [math]::Round($sh.Width,1) + " :: " + $t)
      }
    }
  }
}
Write-Output ("COUNT=" + $n)
Write-Output ("SLIDES=" + $pres.Slides.Count)
$pres.Close(); $app.Quit()

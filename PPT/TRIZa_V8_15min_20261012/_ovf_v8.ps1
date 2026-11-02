$ErrorActionPreference = "Stop"
$path = "D:\doc\ZGCA\Lectures\TRIZ\TRIZa_V8_15min_20261012\TRIZa_V8_AI4Science_15min.pptx"
$app = New-Object -ComObject PowerPoint.Application
$pres = $app.Presentations.Open($path, $true, $false, $false)
$out = @()
foreach ($sl in $pres.Slides) {
  foreach ($sh in $sl.Shapes) {
    if ($sh.HasTextFrame -eq -1) {
      if ($sh.TextFrame.HasText -eq -1) {
        $tr = $sh.TextFrame.TextRange
        $bw = $tr.BoundWidth
        $bh = $tr.BoundHeight
        if ($bh -gt ($sh.Height + 4) -or $bw -gt ($sh.Width + 4)) {
          $txt = $tr.Text -replace "`r", " "
          $txt = $txt -replace "`n", " "
          if ($txt.Length -gt 42) { $txt = $txt.Substring(0, 42) }
          $out += ("S{0} H={1}/{2} W={3}/{4} :: {5}" -f $sl.SlideIndex, [math]::Round($bh,1), [math]::Round($sh.Height,1), [math]::Round($bw,1), [math]::Round($sh.Width,1), $txt)
        }
      }
    }
  }
}
$pres.Close()
$app.Quit()
"COUNT=" + $out.Count
$out

$ErrorActionPreference = "Stop"
$path = "D:\doc\ZGCA\Lectures\TRIZ\FOS_科普_15min_优化版_20261022\FOS_科普_15min_优化版.pptx"
$app = New-Object -ComObject PowerPoint.Application
$pres = $app.Presentations.Open($path, $true, $false, $false)
$sh_h = $pres.PageSetup.SlideHeight
$sw = $pres.PageSetup.SlideWidth
$ovf = New-Object System.Collections.ArrayList
$tot = 0
for ($i = 1; $i -le $pres.Slides.Count; $i++) {
  $sl = $pres.Slides.Item($i)
  for ($j = 1; $j -le $sl.Shapes.Count; $j++) {
    $sh = $sl.Shapes.Item($j)
    if ($sh.HasTextFrame -eq -1) {
      $tf = $sh.TextFrame
      if ($tf.HasText -eq -1) {
        $tot++
        $tr = $tf.TextRange
        $bh = $tr.BoundHeight
        $bw = $tr.BoundWidth
        $hh = $sh.Height
        $ww = $sh.Width
        $bad = $false
        $why = ""
        if ($bh -gt ($hh + 3)) { $bad = $true; $why = "H" }
        if ($bw -gt ($ww + 3)) { $bad = $true; $why = $why + "W" }
        if ($bad) {
          [void]$ovf.Add(("S{0}|{1}|{2}|bh={3:N1}/h={4:N1}|bw={5:N1}/w={6:N1}" -f $i, $sh.Name, $why, $bh, $hh, $bw, $ww))
        }
        $btm = $sh.Top + $sh.Height
        if ($btm -gt ($sh_h + 2)) { [void]$ovf.Add(("S{0}|{1}|OUT-BOTTOM|{2:N1}" -f $i, $sh.Name, $btm)) }
      }
    }
  }
}
$pres.Close()
$app.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
"TEXT_SHAPES=$tot"
"OVERFLOW=$($ovf.Count)"
$ovf | Select-Object -First 40

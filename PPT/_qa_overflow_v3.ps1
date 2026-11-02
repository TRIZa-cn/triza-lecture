$path = "D:\doc\ZGCA\Lectures\TRIZ\FOS_科普_15min_优化版_20261022\FOS_科普_15min_优化版+闰秒案例.pptx"
$tol = 4.0
$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open($path, $true, $false, $false)
$total = 0
$bad = 0
$lines = New-Object System.Collections.ArrayList
foreach ($sl in $pres.Slides) {
  foreach ($sh in $sl.Shapes) {
    if ($sh.HasTextFrame -ne -1) { continue }
    if ($sh.TextFrame.HasText -ne -1) { continue }
    $total++
    $tf = $sh.TextFrame
    $tr = $tf.TextRange
    $aw = $sh.Width - $tf.MarginLeft - $tf.MarginRight
    $ah = $sh.Height - $tf.MarginTop - $tf.MarginBottom
    $bw = $tr.BoundWidth
    $bh = $tr.BoundHeight
    if (($bw -gt ($aw + $tol)) -or ($bh -gt ($ah + $tol))) {
      $bad++
      $t = $tr.Text.Replace("`r", " ").Replace("`n", " ")
      if ($t.Length -gt 30) { $t = $t.Substring(0, 30) }
      [void]$lines.Add(("S{0} W={1:N0}/{2:N0} H={3:N0}/{4:N0} :: {5}" -f $sl.SlideIndex, $bw, $aw, $bh, $ah, $t))
    }
  }
}
$pres.Close()
$ppt.Quit()
"TEXT_SHAPES=$total"
"OVERFLOW=$bad"
$lines | ForEach-Object { $_ }

$path = "D:\doc\ZGCA\Lectures\TRIZ\FOS_科普_15min_优化版_20261022\FOS_科普_15min_优化版+闰秒案例.pptx"
$tol = 1.0
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
    if (($tr.BoundWidth -gt ($aw + $tol)) -or ($tr.BoundHeight -gt ($ah + $tol))) {
      $bad++
      $t = $tr.Text.Replace("`r", " ").Replace("`n", " ")
      if ($t.Length -gt 26) { $t = $t.Substring(0, 26) }
      [void]$lines.Add(("S{0} W={1:N0}/{2:N0} H={3:N0}/{4:N0} :: {5}" -f $sl.SlideIndex, $tr.BoundWidth, $aw, $tr.BoundHeight, $ah, $t))
    }
  }
}
$pres.Close()
$ppt.Quit()
"STRICT_TEXT_SHAPES=$total"
"STRICT_OVERFLOW=$bad"
$lines | ForEach-Object { $_ }

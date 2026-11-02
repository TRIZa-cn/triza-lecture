$path = "D:\doc\ZGCA\Lectures\TRIZ\FOS_科普_15min_优化版_20261022\FOS_科普_15min_优化版+闰秒案例.pptx"
$dir = "D:\doc\ZGCA\Lectures\TRIZ\FOS_科普_15min_优化版_20261022\preview_v3"
if (!(Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open($path, $true, $false, $false)
foreach ($sl in $pres.Slides) {
  $n = "{0:D2}" -f $sl.SlideIndex
  $sl.Export("$dir\p$n.png", "PNG", 1600, 900)
}
$pres.Close()
$ppt.Quit()
(Get-ChildItem $dir -Filter *.png).Count

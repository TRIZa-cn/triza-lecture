param([string]$path)
try {
  $app = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application")
  foreach ($pr in @($app.Presentations)) {
    if ($pr.FullName -eq $path) { $pr.Saved = $true; $pr.Close() ; Write-Output "CLOSED" }
  }
} catch { Write-Output ("NOAPP " + $_.Exception.Message) }

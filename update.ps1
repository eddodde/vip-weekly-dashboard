# update.ps1 - one-command weekly refresh.
# Scans a folder (default: Downloads) for the DRM-protected BI-export xlsx (magic 'SCDSA002'),
# then runs convert.ps1 (Excel COM decrypts DRM on this authorized PC) -> data/perf_long.csv.
# ASCII-only so PowerShell 5.1 cannot corrupt it. Korean filenames are matched by magic bytes,
# not by literal, so no Korean is needed in this script.
#
#   powershell -ExecutionPolicy Bypass -File update.ps1
#   powershell -ExecutionPolicy Bypass -File update.ps1 -SrcDir "C:\some\folder"

param(
  [string]$SrcDir = (Join-Path $env:USERPROFILE "Downloads"),
  [string]$Out = (Join-Path $PSScriptRoot "data\perf_long.csv")
)
$ErrorActionPreference = "Stop"

if (-not (Test-Path $SrcDir)) { Write-Error "Folder not found: $SrcDir"; exit 1 }

$all = Get-ChildItem -Path $SrcDir -Filter *.xlsx -File
$sel = New-Object System.Collections.Generic.List[string]
foreach ($f in $all) {
  try {
    $fs = [System.IO.File]::OpenRead($f.FullName)
    $b = New-Object byte[] 8
    [void]$fs.Read($b, 0, 8)
    $fs.Close()
    if ([System.Text.Encoding]::ASCII.GetString($b) -eq "SCDSA002") { $sel.Add($f.FullName) }
  } catch { }
}
if ($sel.Count -eq 0) {
  Write-Host "No DRM (SCDSA) xlsx found; falling back to ALL *.xlsx in folder."
  foreach ($f in $all) { $sel.Add($f.FullName) }
}

Write-Host ("Selected {0} file(s) from {1}:" -f $sel.Count, $SrcDir)
$sel | ForEach-Object { Write-Host ("  " + (Split-Path $_ -Leaf)) }

& (Join-Path $PSScriptRoot "convert.ps1") -Src $sel -Out $Out
Write-Host "Done. Upload data/perf_long.csv to the app (or git commit & push)."

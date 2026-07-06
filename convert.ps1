# convert.ps1 - Parse LF CRM/VIP weekly BI-export xlsx (wide format) into one tidy long CSV.
# ASCII-only logic (Korean values are passed through from Excel cells, never compared as literals),
# so PowerShell 5.1's ANSI reading of this .ps1 cannot corrupt behavior.
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File convert.ps1
#   powershell -ExecutionPolicy Bypass -File convert.ps1 -Src "C:\path\a.xlsx","C:\path\b.xlsx" -Out "data\perf_long.csv"
#
# Output schema (data/perf_long.csv, UTF-8 BOM):
#   grain,perspective,year,period,period_sort,seg1,seg2,metric,value,source

param(
  [string[]]$Src,
  [string]$Out = "data\perf_long.csv"
)

$ErrorActionPreference = "Stop"

# NOTE: keep this file ASCII-only. PowerShell 5.1 reads a BOM-less .ps1 as ANSI and
# corrupts any non-ASCII (Korean) literal. Pass the (Korean) file paths via -Src from
# the caller instead. Korean *values* read out of Excel cells are fine (COM returns UTF-16).
if (-not $Src -or $Src.Count -eq 0) {
  Write-Error "No -Src files given. Pass xlsx paths: convert.ps1 -Src 'a.xlsx','b.xlsx'"
  exit 1
}

function Is-Number($x) {
  if ($null -eq $x) { return $false }
  return ($x -is [double] -or $x -is [int] -or $x -is [long] -or $x -is [decimal] -or $x -is [single])
}

# Extract a 4-digit year (2000-2100) from a cell that may be a number OR a string; else $null.
function Get-Year($x) {
  if ($null -eq $x) { return $null }
  $s = ([string]$x).Trim()
  if ($s -match "^(20[0-9]{2})") { $iv = [int]$Matches[1]; if ($iv -ge 2000 -and $iv -le 2100) { return $iv } }
  return $null
}

# Classify grain from a period label sample (digit-pattern based, no Korean).
#   day   -> contains '/'            e.g. "1/1"
#   week  -> two digit groups        e.g. "01(wol) 1(jucha)"
#   month -> one digit group         e.g. "1(wol)"
function Get-Grain($label) {
  if ($null -eq $label) { return "unknown" }
  $s = [string]$label
  if ($s -match "/") { return "day" }
  if ($s -match "[0-9]+\D+[0-9]+") { return "week" }
  if ($s -match "[0-9]+") { return "month" }
  return "unknown"
}

# Sortable numeric key. day: yyyymmdd, week: yyyymmww, month: yyyymm00.
function Get-Sort($grain, $year, $label) {
  $s = [string]$label
  switch ($grain) {
    "day"   { if ($s -match "([0-9]{1,2})\s*/\s*([0-9]{1,2})") { return ("{0}{1:D2}{2:D2}" -f $year, [int]$Matches[1], [int]$Matches[2]) } }
    "week"  { if ($s -match "([0-9]{1,2})\D+([0-9]{1,2})") { return ("{0}{1:D2}{2:D2}" -f $year, [int]$Matches[1], [int]$Matches[2]) } }
    "month" { if ($s -match "([0-9]{1,2})") { return ("{0}{1:D2}00" -f $year, [int]$Matches[1]) } }
  }
  return ("{0}9999" -f $year)
}

function Norm-Seg($x) {
  if ($null -eq $x) { return "" }
  $s = ([string]$x).Trim()
  if ($s -eq "-" -or $s -eq "" ) { return "" }
  if ($s -match "TOTAL" -or $s -match "Total") { return "TOTAL" }
  return $s
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$records = New-Object System.Collections.Generic.List[object]

foreach ($file in $Src) {
  if (-not (Test-Path $file)) { Write-Host "SKIP (missing): $file"; continue }
  $leaf = Split-Path $file -Leaf
  Write-Host "Parsing: $leaf"
  $wb = $excel.Workbooks.Open($file, 0, $true)
  $ws = $wb.Worksheets.Item(1)
  $v = $ws.UsedRange.Value2
  $rows = $v.GetLength(0)
  $cols = $v.GetLength(1)

  # value-start col = first column on row 1 holding a 4-digit year marker
  $valStart = 0
  for ($c = 1; $c -le $cols; $c++) {
    if ($null -ne (Get-Year $v.GetValue(1, $c))) { $valStart = $c; break }
  }
  if ($valStart -eq 0) { Write-Host "  WARN: no year marker on row1, skipping"; $wb.Close($false); continue }

  # year forward-fill across value cols
  $yearOf = @{}
  $cur = $null
  for ($c = $valStart; $c -le $cols; $c++) {
    $y = Get-Year $v.GetValue(1, $c)
    if ($null -ne $y) { $cur = $y }
    $yearOf[$c] = $cur
  }

  # period labels on row 3
  $labelOf = @{}
  for ($c = $valStart; $c -le $cols; $c++) { $labelOf[$c] = $v.GetValue(3, $c) }

  # grain from first non-null label
  $grain = "unknown"
  for ($c = $valStart; $c -le $cols; $c++) { if ($null -ne $labelOf[$c]) { $grain = Get-Grain $labelOf[$c]; break } }

  # perspective from label-col count
  $K = $valStart - 1
  $perspective = if ($K -ge 7) { "product" } else { "overall" }

  # iterate data rows: any row (r>=2) with >=1 numeric in value cols
  $curMetric = $null
  $curSeg1 = $null
  for ($r = 2; $r -le $rows; $r++) {
    $hasNum = $false
    for ($c = $valStart; $c -le $cols; $c++) { if (Is-Number $v.GetValue($r, $c)) { $hasNum = $true; break } }
    if (-not $hasNum) { continue }

    $m = $v.GetValue($r, 1)
    if ($null -ne $m) { $curMetric = ([string]$m).Trim() }
    if ($null -eq $curMetric) { continue }

    if ($perspective -eq "product") {
      $s1 = $v.GetValue($r, 2)
      if ($null -ne $s1) { $curSeg1 = Norm-Seg $s1 }
      $seg1 = if ($null -ne $curSeg1) { $curSeg1 } else { "" }
      $seg2 = Norm-Seg ($v.GetValue($r, 3))
    } else {
      $seg1 = Norm-Seg ($v.GetValue($r, 2))
      $seg2 = ""
    }

    for ($c = $valStart; $c -le $cols; $c++) {
      $val = $v.GetValue($r, $c)
      if (-not (Is-Number $val)) { continue }
      $yr = $yearOf[$c]
      $lab = $labelOf[$c]
      if ($null -eq $yr -or $null -eq $lab) { continue }
      $labS = ([string]$lab).Trim()
      $sort = Get-Sort $grain $yr $labS
      $records.Add([pscustomobject]@{
        grain = $grain
        perspective = $perspective
        year = $yr
        period = $labS
        period_sort = $sort
        seg1 = $seg1
        seg2 = $seg2
        metric = $curMetric
        value = $val
        source = $leaf
      })
    }
  }
  $wb.Close($false)
}

$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

$outDir = Split-Path $Out -Parent
if ($outDir -and -not (Test-Path $outDir)) { New-Item -ItemType Directory -Force -Path $outDir | Out-Null }

# Write UTF-8 with BOM so Excel and pandas(utf-8-sig) read Korean correctly.
$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine("grain,perspective,year,period,period_sort,seg1,seg2,metric,value,source")
foreach ($rec in $records) {
  $fields = @($rec.grain, $rec.perspective, $rec.year, $rec.period, $rec.period_sort, $rec.seg1, $rec.seg2, $rec.metric, $rec.value, $rec.source)
  $line = ($fields | ForEach-Object {
    $f = [string]$_
    if ($f -match '[",\r\n]') { '"' + ($f -replace '"', '""') + '"' } else { $f }
  }) -join ","
  [void]$sb.AppendLine($line)
}
$enc = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText((Resolve-Path -LiteralPath (Split-Path $Out -Parent)).Path + "\" + (Split-Path $Out -Leaf), $sb.ToString(), $enc)

Write-Host ("Wrote {0} rows -> {1}" -f $records.Count, $Out)

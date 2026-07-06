# events.ps1 - '2026_가변탭.xlsx'의 '가변탭 확정 스케줄' 그리드에서 날짜->행사 원자료를 추출.
# ASCII-only: 한글 행사 분류는 app.py(Python)에서 처리. 여기선 (year,date,text)만 뽑음.
# 그리드: 연도블록이 가로로 나열 (2024=col21, 2025=col31, 2026=col41).
#   블록 base B: B=연도, B+1=주라벨/트랙, B+2..B+8=7일(MON~SUN) 날짜 serial, 이후 행=행사텍스트.
#
#   powershell -ExecutionPolicy Bypass -File events.ps1 -Src "C:\...\2026_가변탭.xlsx" -Out data\events.csv

param(
  [Parameter(Mandatory = $true)][string]$Src,
  [string]$Out = "data\events.csv",
  [int]$SheetIndex = 3   # '가변탭 확정 스케줄' (한글 리터럴 회피 위해 인덱스로 지정)
)
$ErrorActionPreference = "Stop"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false; $excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($Src, 0, $true)
$ws = $wb.Worksheets.Item($SheetIndex)
$v = $ws.UsedRange.Value2
$rows = $v.GetLength(0)

$bases = @{ 2024 = 21; 2025 = 31; 2026 = 41 }
$recs = New-Object System.Collections.Generic.List[object]

foreach ($yr in 2024, 2025, 2026) {
  $B = $bases[$yr]
  $curDates = $null
  for ($r = 1; $r -le $rows; $r++) {
    # date row?  B+2..B+8 all numeric serials
    $d = @(); $isDate = $true
    for ($k = 2; $k -le 8; $k++) {
      $x = $v.GetValue($r, $B + $k)
      if ($x -is [double] -and $x -gt 40000 -and $x -lt 60000) { $d += [int]$x } else { $isDate = $false; break }
    }
    if ($isDate) { $curDates = $d; continue }
    if ($null -eq $curDates) { continue }
    for ($k = 2; $k -le 8; $k++) {
      $x = $v.GetValue($r, $B + $k)
      if ($null -ne $x) {
        $t = ([string]$x).Trim() -replace '\s+', ' '
        if ($t.Length -gt 2 -and $t -notmatch '^\d+$') {
          $dt = [DateTime]::FromOADate($curDates[$k - 2])
          $recs.Add([pscustomobject]@{ year = $yr; date = $dt.ToString('yyyy-MM-dd'); text = $t })
        }
      }
    }
  }
}
$wb.Close($false); $excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

# dedup (year,date,text)
$seen = @{}; $uniq = New-Object System.Collections.Generic.List[object]
foreach ($r in $recs) { $key = "$($r.year)|$($r.date)|$($r.text)"; if (-not $seen.ContainsKey($key)) { $seen[$key] = $true; $uniq.Add($r) } }

$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine("year,date,text")
foreach ($r in $uniq) {
  $t = $r.text; if ($t -match '[",\r\n]') { $t = '"' + ($t -replace '"', '""') + '"' }
  [void]$sb.AppendLine("$($r.year),$($r.date),$t")
}
$dir = Split-Path $Out -Parent
if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
$enc = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText((Resolve-Path -LiteralPath $dir).Path + "\" + (Split-Path $Out -Leaf), $sb.ToString(), $enc)
Write-Host ("Wrote {0} event rows -> {1}" -f $uniq.Count, $Out)

# events.ps1 - '2026_가변탭.xlsx'의 '가변탭 확정 스케줄' 그리드에서 날짜->행사 원자료를 추출.
# ASCII-only: 한글 행사 분류는 app.py(Python)에서 처리. 여기선 (year,date,text)만 뽑음.
# 그리드: 연도블록이 가로로 나열 (2024=col21, 2025=col31, 2026=col41).
#   블록 base B: B=연도, B+1=주라벨/트랙, B+2..B+8=7일(MON~SUN) 날짜 serial, 이후 행=행사텍스트.
#
#   powershell -ExecutionPolicy Bypass -File events.ps1 -Src "C:\...\2026_가변탭.xlsx" -Out data\events.csv

param(
  [Parameter(Mandatory = $true)][string]$Src,
  [string]$Out = "data\events.csv",
  [int]$SheetIndex = 2   # '가변탭 확정 스케줄 2' = 2026 전체(12월) 확정본. (인덱스로 한글 회피)
)
$ErrorActionPreference = "Stop"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false; $excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($Src, 0, $true)

# 시트별 레이아웃이 다름:
#   시트3 '가변탭 확정 스케줄'   = 과거(2024/2025) 정상, 날짜열 = B+2..B+8
#   시트2 '가변탭 확정 스케줄 2' = 올해(2026) 12월까지 확정, 날짜열 = B+1..B+7
# → 2025는 시트3에서, 2026은 시트2에서 뽑아 합친다.
$configs = @(
  @{ sheet = 3; off = 2; bases = @{ 2025 = 31 } },
  @{ sheet = 2; off = 1; bases = @{ 2026 = 41 } }
)
$recs = New-Object System.Collections.Generic.List[object]

foreach ($cfg in $configs) {
  $ws = $wb.Worksheets.Item([int]$cfg.sheet)
  $v = $ws.UsedRange.Value2
  $rows = $v.GetLength(0)
  $off = [int]$cfg.off
  foreach ($yr in $cfg.bases.Keys) {
    $B = [int]$cfg.bases[$yr]
    $curDates = $null
    for ($r = 1; $r -le $rows; $r++) {
      $d = @(); $isDate = $true
      for ($k = $off; $k -le $off + 6; $k++) {
        $x = $v.GetValue($r, $B + $k)
        if ($x -is [double] -and $x -gt 40000 -and $x -lt 60000) { $d += [int]$x } else { $isDate = $false; break }
      }
      if ($isDate) { $curDates = $d; continue }
      if ($null -eq $curDates) { continue }
      for ($k = $off; $k -le $off + 6; $k++) {
        $x = $v.GetValue($r, $B + $k)
        if ($null -ne $x) {
          $t = ([string]$x).Trim() -replace '\s+', ' '
          if ($t.Length -gt 2 -and $t -notmatch '^\d+$') {
            $dt = [DateTime]::FromOADate($curDates[$k - $off])
            # date 열 연도로 필터(블록 라벨과 실제연도 불일치 방지)
            if ($dt.Year -eq [int]$yr) {
              $recs.Add([pscustomobject]@{ year = $yr; date = $dt.ToString('yyyy-MM-dd'); text = $t })
            }
          }
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

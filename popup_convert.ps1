# popup_convert.ps1 — 온사이트 팝업 실적 export(Overall.xlsx) → data/popup.csv
#   팝업(ETAG1)별 노출·클릭·거래 누적치. 파일에 기간이 없어 '기준일'은 -AsOf(없으면 파일 수정일)로 남긴다.
#   갱신: powershell -ExecutionPolicy Bypass -File popup_convert.ps1 -Src "C:\...\Overall.xlsx" [-AsOf 2026-11-17]
param(
  [Parameter(Mandatory=$true)][string]$Src,
  [string]$AsOf = "",
  [string]$Out = ""
)
$ErrorActionPreference = "Stop"
if(-not $Out){ $Out = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "data\popup.csv" }
if(-not $AsOf){ $AsOf = (Get-Item $Src).LastWriteTime.ToString('yyyy-MM-dd') }
$xl = New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
try {
  $wb = $xl.Workbooks.Open($Src, 0, $true)
  $ws = $wb.Worksheets.Item(1); $ur = $ws.UsedRange; $vals = $ur.Value2
  $rows = $ur.Rows.Count; $cols = $ur.Columns.Count
  $hr=0; for($r=1;$r -le [math]::Min(10,$rows);$r++){ for($c=1;$c -le $cols;$c++){ if(([string]$vals.GetValue($r,$c)).Trim() -eq 'ETAG1'){ $hr=$r; break } }; if($hr){ break } }
  if(-not $hr){ throw "ETAG1 header not found" }
  $hdr=@{}; for($c=1;$c -le $cols;$c++){ $h=([string]$vals.GetValue($hr,$c)).Trim(); if($h){ $hdr[$h]=$c } }
  function Num($x){ if($null -eq $x){0.0}else{ $t=([string]$x) -replace '[^0-9.\-]',''; if($t){[double]$t}else{0.0} } }
  function G($r,$n){ if($hdr.ContainsKey($n)){ Num $vals.GetValue($r,$hdr[$n]) } else { 0.0 } }
  $sb=New-Object System.Text.StringBuilder
  [void]$sb.AppendLine("asof,start,etag,landing,impr,click_uv,click_sv,orders,ord_amt,cust,sale_amt")
  $n=0
  for($r=$hr+1;$r -le $rows;$r++){
    $et=([string]$vals.GetValue($r,$hdr['ETAG1'])).Trim()
    if(-not $et -or $et -notmatch '_(\d{8})_'){ continue }     # 총합계 행 제외
    $st=$Matches[1]; $st="{0}-{1}-{2}" -f $st.Substring(0,4),$st.Substring(4,2),$st.Substring(6,2)
    $ld=''; if($hdr.ContainsKey('ETAG_NM')){ $ld=([string]$vals.GetValue($r,$hdr['ETAG_NM'])).Trim() -replace ',','/' }
    [void]$sb.AppendLine(("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10}" -f $AsOf,$st,$et,$ld,
      [int64](G $r '노출횟수'),[int64](G $r '클릭UV'),[int64](G $r '클릭SV'),[int64](G $r '거래건수'),
      [int64](G $r 'ORD_AMT'),[int64](G $r 'CUST'),[int64](G $r 'SALE_AMT')))
    $n++
  }
  [System.IO.File]::WriteAllText($Out,$sb.ToString(),[System.Text.Encoding]::UTF8)
  Write-Output "saved: $Out ($n rows, asof $AsOf)"
  $wb.Close($false)
} finally { $xl.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)|Out-Null }

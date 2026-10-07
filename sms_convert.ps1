# sms_convert.ps1 — 표(문자) BI export(xlsx) → data/sms_long.csv (주차별 집계, tidy)
#   발송일→대시보드 주차(ISO: 목요일 속한 달·주), 그룹별(EV01/라이브/캠페인전체) 타겟·UV·구매고객·거래액 합산.
#   주간 갱신: powershell -ExecutionPolicy Bypass -File sms_convert.ps1 -Src "C:\...\표(문자) (NN).xlsx"
param(
  [Parameter(Mandatory=$true)][string]$Src,
  [string]$Out = (Join-Path $PSScriptRoot "data\sms_long.csv")
)
$ErrorActionPreference = "Stop"
$xl = New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
try {
  $wb = $xl.Workbooks.Open($Src, 0, $true)
  $ws = $wb.Worksheets.Item(1)
  $ur = $ws.UsedRange
  $vals = $ur.Value2
  $rows = $ur.Rows.Count; $cols = $ur.Columns.Count
  # 헤더 인덱스
  $hdr=@{}; for($c=1;$c -le $cols;$c++){ $h=[string]$vals.GetValue(1,$c); if($h){ $hdr[$h.Trim()]=$c } }
  function Col($n){ $hdr[$n] }
  $iDate=Col '발송일'; $iAF=Col 'AF코드'; $iNm=Col '캠페인명'
  $iT=Col '타겟'; $iU=Col 'UV'; $iB=Col '구매고객'; $iS=Col '거래액'
  $agg=@{}   # key "wk|group|metric" -> sum
  $curd=''
  for($r=2;$r -le $rows;$r++){
    $dv=$vals.GetValue($r,$iDate); if($dv){ $curd=[string]$dv }
    if(-not $curd){ continue }
    $p=$curd -split '/'; if($p.Count -lt 2){ continue }
    $mo=[int]$p[0]; $dy=[int]$p[1]
    $dt=Get-Date "2026-$mo-$dy"
    $thu=$dt.AddDays(3 - (([int]$dt.DayOfWeek + 6) % 7))
    $wn=[int](($thu.Day - 1)/7)+1
    $wk="{0:D2}월 {1}주차" -f $thu.Month,$wn
    $nm=[string]$vals.GetValue($r,$iNm); $af=[string]$vals.GetValue($r,$iAF)
    # ★ VIP 대시보드 — 캠페인명에 'VIP'가 든 건만 VIP 대상이다. MKT_라이브본방/재방 등은
    #   전체 대상 발송이라 제외한다(사용자 확인). 최근본·시크릿·라이브고관여·승급유도 등만 남는다.
    if($nm -notmatch 'VIP'){ continue }
    function Num($x){ if($null -eq $x){0.0}else{ $t=([string]$x) -replace '[^0-9.\-]',''; if($t){[double]$t}else{0.0} } }
    $t=Num $vals.GetValue($r,$iT); $u=Num $vals.GetValue($r,$iU)
    $b=Num $vals.GetValue($r,$iB); $s=Num $vals.GetValue($r,$iS)
    # ★ 그룹은 '캠페인명'으로 가른다. AF코드 EV01은 '최근본(파일럿)'과 'MKT_리텐션_미구매'가
    #   함께 달고 있어(후자는 9/30 300,742명 대량 발송), AF코드로 EV01을 잡으면 리마인드 카드에
    #   미구매 리텐션 블라스트가 섞여 유입률이 1%대로 붕괴한다. 최근본은 반드시 이름으로만.
    $groups=@('캠페인전체')
    if($nm -like '*최근본*'){ $groups += 'EV01' }
    if($nm -like '*라이브*'){ $groups += '라이브' }
    if($nm -like '*미구매*' -or $nm -like '*리텐션*'){ $groups += 'CR미구매' }
    foreach($g in $groups){
      foreach($mv in @(@('타겟',$t),@('UV',$u),@('구매고객',$b),@('거래액',$s))){
        $k="$wk|$g|$($mv[0])"
        if(-not $agg.ContainsKey($k)){ $agg[$k]=0.0 }
        $agg[$k]+=$mv[1]
      }
    }
  }
  $sb=New-Object System.Text.StringBuilder
  [void]$sb.AppendLine("week,group,metric,value")
  foreach($k in ($agg.Keys|Sort-Object)){
    $parts=$k -split '\|'
    [void]$sb.AppendLine("$($parts[0]),$($parts[1]),$($parts[2]),$($agg[$k])")
  }
  [System.IO.File]::WriteAllText($Out,$sb.ToString(),[System.Text.Encoding]::UTF8)
  Write-Output "saved: $Out ($($agg.Count) rows)"
  $wb.Close($false)
} finally { $xl.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)|Out-Null }

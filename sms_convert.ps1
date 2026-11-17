# sms_convert.ps1 — 표(문자) BI export(xlsx) → data/sms_long.csv (주차별 집계, tidy)
#   발송일→대시보드 주차(ISO: 목요일 속한 달·주), 그룹별(EV01/라이브/캠페인전체) 타겟·UV·구매고객·거래액 합산.
#   주간 갱신: powershell -ExecutionPolicy Bypass -File sms_convert.ps1 -Src "C:\...\표(문자) (NN).xlsx"
param(
  [Parameter(Mandatory=$true)][string]$Src,
  [string]$Out = "",
  [string]$DailyOut = ""
)
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
if(-not $Out){ $Out = Join-Path $here "data\sms_long.csv" }
if(-not $DailyOut){ $DailyOut = Join-Path $here "data\sms_daily.csv" }
$xl = New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
try {
  $wb = $xl.Workbooks.Open($Src, 0, $true)
  $ws = $wb.Worksheets.Item(1)
  $ur = $ws.UsedRange
  $vals = $ur.Value2
  $rows = $ur.Rows.Count; $cols = $ur.Columns.Count
  # 헤더 인덱스 — 11/16 export부터 헤더 위에 합계 요약 4행이 붙는다. '발송일'이 있는 행을 찾아 헤더로 쓴다.
  $hr=1; for($r=1;$r -le [math]::Min(10,$rows);$r++){ $hit=$false; for($c=1;$c -le $cols;$c++){ if(([string]$vals.GetValue($r,$c)).Trim() -eq '발송일'){ $hit=$true; break } }; if($hit){ $hr=$r; break } }
  $hdr=@{}; for($c=1;$c -le $cols;$c++){ $h=[string]$vals.GetValue($hr,$c); if($h){ $hdr[$h.Trim()]=$c } }
  function Col($n){ $hdr[$n] }
  $iDate=Col '발송일'; $iAF=Col 'AF코드'; $iNm=Col '캠페인명'; $iNo=Col '캠페인No.'
  $iT=Col '타겟'; $iU=Col 'UV'; $iB=Col '구매고객'; $iS=Col '거래액'
  $agg=@{}   # key "wk|group|metric" -> sum
  # ★ 일자·캠페인 단위 원장(data/sms_daily.csv) — 액션 카드(장바구니·최근 조회 등)가 여기서 자동 산출된다.
  $daily=New-Object System.Collections.Generic.List[string]; $dDates=@{}; $minDt=$null
  $nowY=(Get-Date).Year; $nowM=(Get-Date).Month
  function WkOf($d){ $thu=$d.AddDays(3 - (([int]$d.DayOfWeek + 6) % 7)); $wn=[int][math]::Floor(($thu.Day - 1)/7)+1; "{0:D2}월 {1}주차" -f $thu.Month,$wn }
  $curd=''
  for($r=$hr+1;$r -le $rows;$r++){
    $dv=$vals.GetValue($r,$iDate); if($dv){ $curd=[string]$dv }
    if(-not $curd){ continue }
    $p=$curd -split '/'; if($p.Count -lt 2){ continue }
    $mo=[int]$p[0]; $dy=[int]$p[1]
    $yr=$nowY; if($mo -gt $nowM + 1){ $yr=$nowY-1 }   # 연초에 전년 12월 export를 돌리는 경우
    $dt=Get-Date -Year $yr -Month $mo -Day $dy -Hour 0 -Minute 0 -Second 0
    if($null -eq $minDt -or $dt -lt $minDt){ $minDt=$dt }
    $dIso=$dt.ToString('yyyy-MM-dd'); $dDates[$dIso]=1
    $wk=WkOf $dt
    $nm=[string]$vals.GetValue($r,$iNm); $af=[string]$vals.GetValue($r,$iAF)
    # ★ VIP 대시보드 — 캠페인명에 'VIP'가 든 건만 VIP 대상이다. MKT_라이브본방/재방 등은
    #   전체 대상 발송이라 제외한다(사용자 확인). 최근본·시크릿·라이브고관여·승급유도 등만 남는다.
    if($nm -notmatch 'VIP'){ continue }
    function Num($x){ if($null -eq $x){0.0}else{ $t=([string]$x) -replace '[^0-9.\-]',''; if($t){[double]$t}else{0.0} } }
    $t=Num $vals.GetValue($r,$iT); $u=Num $vals.GetValue($r,$iU)
    $b=Num $vals.GetValue($r,$iB); $s=Num $vals.GetValue($r,$iS)
    $no=''; if($iNo){ $no=([string]$vals.GetValue($r,$iNo)).Trim() }
    $daily.Add(("{0},{1},{2},{3},{4},{5},{6},`"{7}`"" -f $dIso,$no,$af.Trim(),[int64]$t,[int64]$u,[int64]$b,[int64]$s,($nm.Trim() -replace '"','""')))
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
  # ★ export 첫 날짜가 월요일이 아니면 그 주는 앞부분이 잘린 주다. 이미 저장된 주차라면 덮어쓰지 않는다
  #   (예: 11/1(일)부터 받은 export가 10월 5주차를 하루치로 바꿔 버리는 것 방지).
  if($minDt -and [int]$minDt.DayOfWeek -ne 1 -and (Test-Path $Out)){
    $lw=WkOf $minDt; $has=$false
    foreach($ln in (Get-Content $Out -Encoding UTF8 | Select-Object -Skip 1)){ if($ln -and (($ln -split ',')[0] -eq $lw)){ $has=$true; break } }
    if($has){ foreach($k in @($agg.Keys)){ if(($k -split '\|')[0] -eq $lw){ $agg.Remove($k) } }; Write-Output "skip partial leading week: $lw" }
  }
  # 기존 파일의 다른 주차는 보존하고, 이번 export에 든 주차만 교체한다(export가 1주 단위로 와도 누적 유지).
  $newWk=@{}; foreach($k in $agg.Keys){ $newWk[($k -split '\|')[0]]=1 }
  $sb=New-Object System.Text.StringBuilder
  [void]$sb.AppendLine("week,group,metric,value")
  if(Test-Path $Out){
    foreach($ln in (Get-Content $Out -Encoding UTF8 | Select-Object -Skip 1)){
      if(-not $ln){ continue }
      if(-not $newWk.ContainsKey(($ln -split ',')[0])){ [void]$sb.AppendLine($ln) }
    }
  }
  foreach($k in ($agg.Keys|Sort-Object)){
    $parts=$k -split '\|'
    [void]$sb.AppendLine("$($parts[0]),$($parts[1]),$($parts[2]),$($agg[$k])")
  }
  [System.IO.File]::WriteAllText($Out,$sb.ToString(),[System.Text.Encoding]::UTF8)
  Write-Output "saved: $Out ($($agg.Count) rows)"
  # 일자 원장: export에 든 날짜만 교체, 나머지 날짜는 보존
  $sd=New-Object System.Text.StringBuilder
  [void]$sd.AppendLine("date,camp_no,af,target,uv,buyers,sales,name")
  $keep=New-Object System.Collections.Generic.List[string]
  if(Test-Path $DailyOut){
    foreach($ln in (Get-Content $DailyOut -Encoding UTF8 | Select-Object -Skip 1)){
      if($ln -and -not $dDates.ContainsKey(($ln -split ',')[0])){ $keep.Add($ln) }
    }
  }
  foreach($ln in (@($keep)+@($daily) | Sort-Object)){ [void]$sd.AppendLine($ln) }
  [System.IO.File]::WriteAllText($DailyOut,$sd.ToString(),[System.Text.Encoding]::UTF8)
  Write-Output "saved: $DailyOut ($($daily.Count) new rows, $($keep.Count) kept)"
  $wb.Close($false)
} finally { $xl.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)|Out-Null }

param([string]$OutDir,[ValidateSet('Quick','Deep')][string]$Mode='Quick')
$ErrorActionPreference='Continue'
$results=New-Object System.Collections.ArrayList
function Add-Result([string]$Name,[string]$Status,[int]$Score,[string]$Detail,[int]$Weight){
 [void]$results.Add([pscustomobject]@{Name=$Name;Status=$Status;Score=$Score;Detail=$Detail;Weight=$Weight})
}
function Stage([int]$Percent,[string]$Message){[pscustomobject]@{Percent=$Percent;Message=$Message}|ConvertTo-Json -Compress|Set-Content -LiteralPath (Join-Path $OutDir 'progress.json') -Encoding UTF8}
function Html([string]$s){[System.Net.WebUtility]::HtmlEncode($s)}
try{
 Stage 5 'Checking battery...'
 try{
  $battery=Get-CimInstance Win32_Battery -ErrorAction Stop | Select-Object -First 1
  $design=Get-CimInstance -Namespace root/wmi -Class BatteryStaticData -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty DesignedCapacity
  $full=Get-CimInstance -Namespace root/wmi -Class BatteryFullChargedCapacity -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullChargedCapacity
  if($design -gt 0 -and $full -gt 0){$pct=[math]::Min(100,[math]::Round(100*$full/$design));$status=if($pct -ge 80){'PASS'}elseif($pct -ge 65){'WARN'}else{'FAIL'};Add-Result 'Battery' $status $pct "Capacity $full / $design mWh ($pct%). Runtime not measured." 25}
  elseif($battery){Add-Result 'Battery' 'UNKNOWN' -1 'Battery detected, but capacity readings unavailable.' 25}
  else{Add-Result 'Battery' 'UNKNOWN' -1 'No readable battery (desktop or unsupported firmware).' 25}
 }catch{Add-Result 'Battery' 'UNKNOWN' -1 $_.Exception.Message 25}
 Stage 20 'Checking storage health...'
 try{
  $disks=@(Get-PhysicalDisk -ErrorAction Stop)
  if($disks.Count -eq 0){Add-Result 'SSD / HDD' 'UNKNOWN' -1 'Storage device health not exposed.' 25}
  else{
   $bad=@($disks|Where-Object {$_.HealthStatus -eq 'Unhealthy'});$warn=@($disks|Where-Object {$_.HealthStatus -eq 'Warning'});
   $status=if($bad.Count){'FAIL'}elseif($warn.Count){'WARN'}elseif(@($disks|Where-Object {$_.HealthStatus -eq 'Healthy'}).Count -eq $disks.Count){'PASS'}else{'UNKNOWN'}
   $score=if($status -eq 'PASS'){95}elseif($status -eq 'WARN'){55}elseif($status -eq 'FAIL'){10}else{-1}
   $detail=($disks|ForEach-Object {"$($_.FriendlyName): $($_.HealthStatus), $([math]::Round($_.Size/1GB)) GB"}) -join '; '
   try{$counters=@(Get-StorageReliabilityCounter -PhysicalDisk $disks -ErrorAction Stop);$extras=@($counters|ForEach-Object {if($_.Wear -ne $null){"Wear $($_.Wear)%"};if($_.Temperature -gt 0){"Temp $($_.Temperature) C"}});if($extras.Count){$detail+='; '+($extras -join ', ')}}catch{}
   Add-Result 'SSD / HDD' $status $score $detail 25
  }
 }catch{Add-Result 'SSD / HDD' 'UNKNOWN' -1 'Disk health unavailable: '+$_.Exception.Message 25}
 Stage 38 'Checking RAM...'
 try{
  $os=Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
  $modules=@(Get-CimInstance Win32_PhysicalMemory -ErrorAction Stop)
  $total=[math]::Round($os.TotalVisibleMemorySize/1MB,1)
  $size=if($Mode -eq 'Deep'){128MB}else{32MB}
  $buf=New-Object byte[] $size
  for($i=0;$i -lt $buf.Length;$i+=4096){$buf[$i]=[byte](($i/4096)%251)}
  $ok=$true;for($i=0;$i -lt $buf.Length;$i+=4096){if($buf[$i] -ne [byte](($i/4096)%251)){$ok=$false;break}}
  $buf=$null
  $status=if($ok){'PASS'}else{'FAIL'};$score=if($ok){90}else{0}
  Add-Result 'RAM' $status $score "$total GB, $($modules.Count) module(s). Small allocation check only; not a full RAM test." 15
 }catch{Add-Result 'RAM' 'UNKNOWN' -1 $_.Exception.Message 15}
 Stage 55 'Checking CPU...'
 try{
  $cpu=Get-CimInstance Win32_Processor -ErrorAction Stop|Select-Object -First 1
  $loops=if($Mode -eq 'Deep'){1800000}else{350000}
  $sw=[diagnostics.stopwatch]::StartNew();$v=0.0;for($i=1;$i -le $loops;$i++){$v+=[math]::Sqrt($i)};$sw.Stop()
  $ok=([double]::IsNaN($v) -eq $false -and $v -gt 0)
  Add-Result 'CPU' $(if($ok){'PASS'}else{'FAIL'}) $(if($ok){90}else{0}) "$($cpu.Name.Trim()); $($cpu.NumberOfCores) cores; computation check $([math]::Round($sw.Elapsed.TotalSeconds,2))s. No thermal stress validation." 15
 }catch{Add-Result 'CPU' 'UNKNOWN' -1 $_.Exception.Message 15}
 Stage 75 'Checking GPU, network, and system...'
 try{
  $gpu=@(Get-CimInstance Win32_VideoController -ErrorAction Stop)
  if($gpu.Count){$bad=@($gpu|Where-Object {$_.ConfigManagerErrorCode -ne 0});Add-Result 'GPU / Display adapter' $(if($bad.Count){'WARN'}else{'PASS'}) $(if($bad.Count){50}else{90}) (($gpu|ForEach-Object {"$($_.Name) (device code $($_.ConfigManagerErrorCode))"}) -join '; ') 10}
  else{Add-Result 'GPU / Display adapter' 'UNKNOWN' -1 'No adapter data.' 10}
 }catch{Add-Result 'GPU / Display adapter' 'UNKNOWN' -1 $_.Exception.Message 10}
 try{
  $adapters=@(Get-NetAdapter -ErrorAction Stop|Where-Object {$_.InterfaceDescription -match 'Wireless|Wi-Fi|802.11'})
  if($adapters.Count){$up=@($adapters|Where-Object {$_.Status -eq 'Up'});Add-Result 'Wi-Fi adapter' $(if($up.Count){'PASS'}else{'WARN'}) $(if($up.Count){95}else{60}) (($adapters|ForEach-Object {"$($_.Name): $($_.Status)"}) -join '; ') 5}
  else{Add-Result 'Wi-Fi adapter' 'UNKNOWN' -1 'No identifiable wireless adapter.' 5}
 }catch{Add-Result 'Wi-Fi adapter' 'UNKNOWN' -1 $_.Exception.Message 5}
 try{
  $os=Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
  $sys=Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
  Add-Result 'Windows / System' 'PASS' 90 "$($sys.Manufacturer) $($sys.Model); $($os.Caption); last boot $($os.LastBootUpTime)" 5
 }catch{Add-Result 'Windows / System' 'UNKNOWN' -1 $_.Exception.Message 5}
 Stage 90 'Generating report...'
 $known=@($results|Where-Object {$_.Score -ge 0});$sum=0;$weights=0
 foreach($r in $known){$sum+=$r.Score*$r.Weight;$weights+=$r.Weight}
 $overall=if($weights){[math]::Round($sum/$weights)}else{-1}
 $coverage=[math]::Round($weights)
 $critical=@($results|Where-Object {$_.Status -eq 'FAIL' -and $_.Name -in @('Battery','SSD / HDD','RAM','CPU')})
 $verdict=if($critical.Count){'AVOID / INVESTIGATE'}elseif($coverage -lt 70){'INSUFFICIENT DATA'}elseif(@($results|Where-Object {$_.Status -eq 'WARN'}).Count){'INSPECT FURTHER'}elseif($overall -ge 80){'PROMISING - VERIFY MANUALLY'}else{'INSPECT FURTHER'}
 $stamp=Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
 $report=[pscustomobject]@{Mode=$Mode;Date=$stamp;Overall=$overall;Coverage=$coverage;Verdict=$verdict;Results=@($results)}
 $report|ConvertTo-Json -Depth 6|Set-Content (Join-Path $OutDir 'report.json') -Encoding UTF8
 $rows=($results|ForEach-Object {"<tr><td>$(Html $_.Name)</td><td class='$($_.Status)'>$(Html $_.Status)</td><td>$(if($_.Score -lt 0){'N/A'}else{"$($_.Score)/100"})</td><td>$(Html $_.Detail)</td></tr>"}) -join "`n"
 $html=@"
<!doctype html><html lang="en"><head><meta charset="utf-8"><title>Laptop Health Inspection</title><style>body{font:15px Segoe UI,Arial;background:#f4f7fb;color:#17243b;margin:35px auto;max-width:1050px;padding:0 20px}header{background:#14243c;color:white;border-radius:16px;padding:28px}h1{margin:0 0 8px}.score{font-size:48px;font-weight:800}table{border-collapse:collapse;width:100%;background:white;margin-top:22px}th,td{text-align:left;border-bottom:1px solid #e5e9ef;padding:13px;vertical-align:top}th{background:#e9eef5}.PASS{color:#13824b;font-weight:bold}.WARN{color:#ae7200;font-weight:bold}.FAIL{color:#c12c39;font-weight:bold}.UNKNOWN{color:#67758a;font-weight:bold}.note{background:#fff4d8;padding:18px;border-radius:12px;margin-top:22px;line-height:1.6}</style></head><body><header><h1>Laptop Health Checker</h1><div>$(Html $stamp) | $(Html $Mode) inspection</div><div class="score">$(if($overall -lt 0){'N/A'}else{"$overall/100"})</div><strong>$(Html $verdict)</strong><div>Diagnostic coverage: $coverage% of weighted categories</div></header><table><thead><tr><th>Component</th><th>Status</th><th>Inspection score</th><th>Details</th></tr></thead><tbody>$rows</tbody></table><div class="note"><strong>Important:</strong> Scores are heuristic inspection ratings, not measured remaining lifespan or probability of reliability. Battery percentage estimates remaining full-charge capacity, not runtime. SMART and firmware may hide disk defects. RAM allocation is not a full memory test. CPU/GPU tests are not long thermal or graphics stress tests. Manually inspect screen, keyboard, hinges, speakers, ports, webcam, charging, battery runtime and seller ownership before buying.</div></body></html>
"@
 $html|Set-Content (Join-Path $OutDir 'report.html') -Encoding UTF8
 Stage 100 'Completed'
}catch{
 $_|Out-String|Set-Content (Join-Path $OutDir 'error.txt') -Encoding UTF8
 Stage 100 'Error: inspection incomplete'
}

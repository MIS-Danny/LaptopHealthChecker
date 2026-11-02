param(
    [Parameter(Mandatory=$true)][string]$OutDir,
    [ValidateSet('Quick','Deep')][string]$Mode='Quick'
)
$ErrorActionPreference='Continue'
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$results=New-Object System.Collections.ArrayList
function Add-Result([string]$Name,[string]$Status,[int]$Score,[string]$Detail,[int]$Weight){
    [void]$results.Add([pscustomobject]@{Name=$Name;Status=$Status;Score=$Score;Detail=$Detail;Weight=$Weight})
}
function Stage([int]$Percent,[string]$Message){
    [pscustomobject]@{Percent=$Percent;Message=$Message} | ConvertTo-Json -Compress | Set-Content -LiteralPath (Join-Path $OutDir 'progress.json') -Encoding UTF8
}
function Html([string]$s){[System.Net.WebUtility]::HtmlEncode($s)}
function Read-BatteryReport {
    $path=Join-Path $OutDir 'battery-report.html'
    $exe=Join-Path $env:SystemRoot 'System32\powercfg.exe'
    & $exe /batteryreport /output $path 2>&1 | Out-Null
    if(!(Test-Path -LiteralPath $path)){return $null}
    $html=Get-Content -LiteralPath $path -Raw -ErrorAction Stop
    # Parse only the Installed batteries section, to avoid values in historical tables.
    $start=[regex]::Match($html,'Installed batteries', 'IgnoreCase')
    if(!$start.Success){return $null}
    $part=$html.Substring($start.Index,[math]::Min(35000,$html.Length-$start.Index))
    $next=[regex]::Match($part,'Recent usage|Battery usage|Usage history','IgnoreCase')
    if($next.Success){$part=$part.Substring(0,$next.Index)}
    $plain=[System.Net.WebUtility]::HtmlDecode(($part -replace '(?is)<script\b[^>]*>.*?</script>',' ' -replace '(?is)<style\b[^>]*>.*?</style>',' ' -replace '(?i)</(?:tr|td|th|div|p|span|h\d)>',"`n" -replace '(?s)<[^>]+>',' '))
    $plain=$plain -replace '\u00a0',' '
    $designMatch=[regex]::Match($plain,'DESIGN\s+CAPACITY\s*([\d,\.]+)\s*mWh','IgnoreCase')
    $fullMatch=[regex]::Match($plain,'FULL\s+CHARGE\s+CAPACITY\s*([\d,\.]+)\s*mWh','IgnoreCase')
    $cycleMatch=[regex]::Match($plain,'CYCLE\s+COUNT\s*([\d,]+|-)','IgnoreCase')
    if(!$designMatch.Success -or !$fullMatch.Success){return $null}
    $design=0.0;$full=0.0
    if(![double]::TryParse(($designMatch.Groups[1].Value -replace ',',''),[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$design)){return $null}
    if(![double]::TryParse(($fullMatch.Groups[1].Value -replace ',',''),[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$full)){return $null}
    if($design -le 0 -or $full -le 0){return $null}
    $cycles=if($cycleMatch.Success){$cycleMatch.Groups[1].Value}else{'Unavailable'}
    return [pscustomobject]@{Design=$design;Full=$full;Cycles=$cycles}
}
try {
    Stage 5 'Checking battery...'
    try {
        $battery=Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue | Select-Object -First 1
        $design=Get-CimInstance -Namespace root/wmi -Class BatteryStaticData -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty DesignedCapacity
        $full=Get-CimInstance -Namespace root/wmi -Class BatteryFullChargedCapacity -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullChargedCapacity
        $source='WMI';$cycles='Unavailable'
        if(!($design -gt 0 -and $full -gt 0)){
            $fallback=Read-BatteryReport
            if($fallback){$design=$fallback.Design;$full=$fallback.Full;$cycles=$fallback.Cycles;$source='Windows battery report'}
        }
        if($design -gt 0 -and $full -gt 0){
            $pct=[math]::Round(100*[double]$full/[double]$design,1)
            $score=[int][math]::Min(100,[math]::Round($pct))
            $state=if($pct -ge 80){'PASS'}elseif($pct -ge 65){'WARN'}else{'FAIL'}
            $detail="Health $pct%; design $([math]::Round($design)) mWh; full charge $([math]::Round($full)) mWh; cycles $cycles; source $source. Runtime not measured."
            Add-Result 'Battery' $state $score $detail 25
        } elseif($battery){
            Add-Result 'Battery' 'UNKNOWN' -1 'Battery detected, but capacity data unavailable from WMI and Windows battery report.' 25
        } else {
            Add-Result 'Battery' 'UNKNOWN' -1 'Battery not detected or capacity information unavailable.' 25
        }
    }catch{Add-Result 'Battery' 'UNKNOWN' -1 $_.Exception.Message 25}
    Stage 20 'Checking storage health...'
    try {
        $disks=@(Get-PhysicalDisk -ErrorAction Stop)
        if($disks.Count -eq 0){Add-Result 'SSD / HDD' 'UNKNOWN' -1 'Storage device health not exposed.' 25}
        else {
            $bad=@($disks|Where-Object {$_.HealthStatus -eq 'Unhealthy'})
            $warn=@($disks|Where-Object {$_.HealthStatus -eq 'Warning'})
            $state=if($bad.Count){'FAIL'}elseif($warn.Count){'WARN'}elseif(@($disks|Where-Object {$_.HealthStatus -eq 'Healthy'}).Count -eq $disks.Count){'PASS'}else{'UNKNOWN'}
            $diskScore=if($state -eq 'PASS'){95}elseif($state -eq 'WARN'){55}elseif($state -eq 'FAIL'){10}else{-1}
            $detail=($disks|ForEach-Object {"$($_.FriendlyName): $($_.HealthStatus), $([math]::Round($_.Size/1GB)) GB"}) -join '; '
            try {
                $counters=@(Get-StorageReliabilityCounter -PhysicalDisk $disks -ErrorAction Stop)
                $extras=@($counters|ForEach-Object {if($null -ne $_.Wear){"Wear $($_.Wear)%"};if($_.Temperature -gt 0){"Temp $($_.Temperature) C"}})
                if($extras.Count){$detail+='; '+($extras -join ', ')}
            }catch{}
            Add-Result 'SSD / HDD' $state $diskScore $detail 25
        }
    }catch{Add-Result 'SSD / HDD' 'UNKNOWN' -1 ('Disk health unavailable: '+$_.Exception.Message) 25}
    Stage 38 'Checking RAM...'
    try {
        $os=Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        $modules=@(Get-CimInstance Win32_PhysicalMemory -ErrorAction Stop)
        $total=[math]::Round($os.TotalVisibleMemorySize/1MB,1)
        $size=if($Mode -eq 'Deep'){128MB}else{32MB}
        $buf=New-Object byte[] $size
        for($i=0;$i -lt $buf.Length;$i+=4096){$buf[$i]=[byte](($i/4096)%251)}
        $ok=$true
        for($i=0;$i -lt $buf.Length;$i+=4096){if($buf[$i] -ne [byte](($i/4096)%251)){$ok=$false;break}}
        $buf=$null
        Add-Result 'RAM' $(if($ok){'PASS'}else{'FAIL'}) $(if($ok){90}else{0}) "$total GB, $($modules.Count) module(s). Small allocation check only; not a full RAM test." 15
    }catch{Add-Result 'RAM' 'UNKNOWN' -1 $_.Exception.Message 15}
    Stage 55 'Checking CPU...'
    try {
        $cpu=Get-CimInstance Win32_Processor -ErrorAction Stop|Select-Object -First 1
        $loops=if($Mode -eq 'Deep'){1800000}else{350000}
        $sw=[diagnostics.stopwatch]::StartNew();$v=0.0
        for($i=1;$i -le $loops;$i++){$v+=[math]::Sqrt($i)}
        $sw.Stop()
        $ok=([double]::IsNaN($v) -eq $false -and $v -gt 0)
        Add-Result 'CPU' $(if($ok){'PASS'}else{'FAIL'}) $(if($ok){90}else{0}) "$($cpu.Name.Trim()); $($cpu.NumberOfCores) cores; computation check $([math]::Round($sw.Elapsed.TotalSeconds,2))s. No thermal stress validation." 15
    }catch{Add-Result 'CPU' 'UNKNOWN' -1 $_.Exception.Message 15}
    Stage 75 'Checking GPU, network, and system...'
    try {
        $gpu=@(Get-CimInstance Win32_VideoController -ErrorAction Stop)
        if($gpu.Count){
            $bad=@($gpu|Where-Object {$_.ConfigManagerErrorCode -ne 0})
            Add-Result 'GPU / Display adapter' $(if($bad.Count){'WARN'}else{'PASS'}) $(if($bad.Count){50}else{90}) (($gpu|ForEach-Object {"$($_.Name) (device code $($_.ConfigManagerErrorCode))"}) -join '; ') 10
        }else{Add-Result 'GPU / Display adapter' 'UNKNOWN' -1 'No adapter data.' 10}
    }catch{Add-Result 'GPU / Display adapter' 'UNKNOWN' -1 $_.Exception.Message 10}
    try {
        $adapters=@(Get-NetAdapter -ErrorAction Stop|Where-Object {$_.InterfaceDescription -match 'Wireless|Wi-Fi|802.11'})
        if($adapters.Count){
            $up=@($adapters|Where-Object {$_.Status -eq 'Up'})
            $disabled=@($adapters|Where-Object {$_.Status -eq 'Disabled'})
            $wifiState=if($up.Count){'PASS'}elseif($disabled.Count -eq $adapters.Count){'WARN'}else{'UNKNOWN'}
            $wifiScore=if($wifiState -eq 'PASS'){95}elseif($wifiState -eq 'WARN'){60}else{-1}
            Add-Result 'Wi-Fi adapter' $wifiState $wifiScore (($adapters|ForEach-Object {"$($_.Name): $($_.Status)"}) -join '; ') 5
        }else{Add-Result 'Wi-Fi adapter' 'UNKNOWN' -1 'No identifiable wireless adapter.' 5}
    }catch{Add-Result 'Wi-Fi adapter' 'UNKNOWN' -1 $_.Exception.Message 5}
    try {
        $os=Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        $sys=Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
        Add-Result 'Windows / System' 'PASS' 90 "$($sys.Manufacturer) $($sys.Model); $($os.Caption); last boot $($os.LastBootUpTime)" 5
    }catch{Add-Result 'Windows / System' 'UNKNOWN' -1 $_.Exception.Message 5}
    # Supplemental automatic checks are informational and do not inflate the weighted health score.
    Stage 82 'Checking additional system diagnostics...'
    try {
        $volumes=@(Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' -ErrorAction Stop)
        if($volumes.Count){
            $low=@($volumes | Where-Object {$_.Size -gt 0 -and ($_.FreeSpace / $_.Size) -lt 0.10})
            $detail=($volumes | ForEach-Object {if($_.Size -gt 0){"$($_.DeviceID) $([math]::Round($_.FreeSpace/1GB,1)) GB free of $([math]::Round($_.Size/1GB,1)) GB"}}) -join '; '
            Add-Result 'Storage free space' $(if($low.Count){'WARN'}else{'PASS'}) $(if($low.Count){60}else{95}) $detail 0
        }else{Add-Result 'Storage free space' 'UNKNOWN' -1 'No fixed-volume capacity data.' 0}
    }catch{Add-Result 'Storage free space' 'UNKNOWN' -1 $_.Exception.Message 0}
    try {
        $modules=@(Get-CimInstance Win32_PhysicalMemory -ErrorAction Stop)
        if($modules.Count){
            $detail=($modules | ForEach-Object {"$([math]::Round($_.Capacity/1GB,1)) GB @ $($_.Speed) MHz ($($_.Manufacturer))"}) -join '; '
            Add-Result 'RAM configuration' 'PASS' 90 "$($modules.Count) module(s); $detail. Configuration only, not memory integrity." 0
        }else{Add-Result 'RAM configuration' 'UNKNOWN' -1 'No memory module data.' 0}
    }catch{Add-Result 'RAM configuration' 'UNKNOWN' -1 $_.Exception.Message 0}
    try {
        $bad=@(Get-CimInstance Win32_PnPEntity -Filter 'ConfigManagerErrorCode <> 0' -ErrorAction Stop | Where-Object {$_.ConfigManagerErrorCode -notin @(0,22,24,45)})
        $detail=if($bad.Count){($bad | Select-Object -First 8 | ForEach-Object {"$($_.Name) (code $($_.ConfigManagerErrorCode))"}) -join '; '}else{'No active device error codes reported. Disabled/disconnected devices excluded.'}
        Add-Result 'Device Manager errors' $(if($bad.Count){'WARN'}else{'PASS'}) $(if($bad.Count){55}else{95}) $detail 0
    }catch{Add-Result 'Device Manager errors' 'UNKNOWN' -1 $_.Exception.Message 0}
    try {
        $bios=Get-CimInstance Win32_BIOS -ErrorAction Stop | Select-Object -First 1
        Add-Result 'BIOS / firmware info' 'PASS' 90 "Version $($bios.SMBIOSBIOSVersion); release $($bios.ReleaseDate). Informational only; not a firmware security audit." 0
    }catch{Add-Result 'BIOS / firmware info' 'UNKNOWN' -1 $_.Exception.Message 0}
    try {
        $os=Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        $uptime=(Get-Date)-$os.LastBootUpTime
        Add-Result 'System uptime' 'PASS' 90 "Uptime $([math]::Round($uptime.TotalHours,1)) hours; last boot $($os.LastBootUpTime). Informational only." 0
    }catch{Add-Result 'System uptime' 'UNKNOWN' -1 $_.Exception.Message 0}
    try {
        $pending=(Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') -or (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired')
        Add-Result 'Pending Windows reboot' $(if($pending){'WARN'}else{'PASS'}) $(if($pending){60}else{95}) $(if($pending){'Windows indicates a pending reboot. Not necessarily a hardware problem.'}else{'No reboot pending in checked Windows Update/servicing registry locations.'}) 0
    }catch{Add-Result 'Pending Windows reboot' 'UNKNOWN' -1 $_.Exception.Message 0}
    try {
        $since=(Get-Date).AddDays(-7)
        $events=@(Get-WinEvent -FilterHashtable @{LogName='System';Level=1;StartTime=$since} -MaxEvents 10 -ErrorAction Stop)
        $detail=if($events.Count){($events | Select-Object -First 5 | ForEach-Object {"$($_.TimeCreated.ToString('yyyy-MM-dd')) ID $($_.Id) $($_.ProviderName)"}) -join '; '}else{'No critical System events found in last 7 days.'}
        Add-Result 'Recent critical events' $(if($events.Count){'WARN'}else{'PASS'}) $(if($events.Count){60}else{95}) $detail 0
    }catch{Add-Result 'Recent critical events' 'UNKNOWN' -1 'Event log unavailable or access denied.' 0}
    Stage 90 'Generating report...'
    $known=@($results|Where-Object {$_.Score -ge 0});$sum=0;$weights=0
    foreach($r in $known){$sum+=$r.Score*$r.Weight;$weights+=$r.Weight}
    $overall=if($weights){[math]::Round($sum/$weights)}else{-1}
    $totalWeights=($results|Measure-Object -Property Weight -Sum).Sum
    $coverage=if($totalWeights){[math]::Round(100*$weights/$totalWeights)}else{0}
    # Buying recommendation: heuristic screening, never a guarantee of safety.
    $critical=@($results|Where-Object {$_.Status -eq 'FAIL' -and $_.Name -in @('Battery','SSD / HDD','RAM','CPU')})
    $batteryUnknown=@($results|Where-Object {$_.Name -eq 'Battery' -and $_.Status -eq 'UNKNOWN'}).Count -gt 0
    $warnings=@($results|Where-Object {$_.Status -eq 'WARN'})
    $verdict=if($critical.Count){'DO NOT BUY YET - CRITICAL FAILURE'}elseif($coverage -lt 90){'INCOMPLETE TESTING'}elseif($batteryUnknown){'CAUTION - BATTERY UNVERIFIED'}elseif($warnings.Count){'CAUTION - CHECK WARNINGS'}elseif($overall -ge 85){'PROMISING - PHYSICAL CHECK REQUIRED'}elseif($overall -ge 70){'CAUTION - INSPECT CAREFULLY'}else{'HIGH RISK - NOT RECOMMENDED'}
    $stamp=Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $report=[pscustomobject]@{Mode=$Mode;Date=$stamp;Overall=$overall;Coverage=$coverage;Verdict=$verdict;Results=@($results)}
    $report|ConvertTo-Json -Depth 6|Set-Content (Join-Path $OutDir 'report.json') -Encoding UTF8
    $rows=($results|ForEach-Object {"<tr><td>$(Html $_.Name)</td><td class='$($_.Status)'>$(Html $_.Status)</td><td>$(if($_.Score -lt 0){'N/A'}else{"$($_.Score)/100"})</td><td>$(Html $_.Detail)</td></tr>"}) -join "`n"
    $html=@"
<!doctype html><html lang="en"><head><meta charset="utf-8"><title>Laptop Health Inspection</title><style>body{font:15px Segoe UI,Arial;background:#0b1220;color:#e2e8f0;margin:35px auto;max-width:1050px;padding:0 20px}header{background:#172337;border-radius:16px;padding:28px}h1{margin:0 0 8px}.score{font-size:48px;font-weight:800}table{border-collapse:collapse;width:100%;background:#172337;margin-top:22px}th,td{text-align:left;border-bottom:1px solid #334155;padding:13px;vertical-align:top}th{background:#24334b}.PASS{color:#4ade80;font-weight:bold}.WARN{color:#fbbf24;font-weight:bold}.FAIL{color:#fb7185;font-weight:bold}.UNKNOWN{color:#94a3b8;font-weight:bold}.note{background:#24334b;padding:18px;border-radius:12px;margin-top:22px;line-height:1.6}</style></head><body><header><h1>Laptop Health Checker</h1><div>$(Html $stamp) | $(Html $Mode) inspection</div><div class="score">$(if($overall -lt 0){'N/A'}else{"$overall/100"})</div><strong>Buying recommendation: $(Html $verdict)</strong><div>Diagnostic coverage: $coverage% of weighted categories</div></header><table><thead><tr><th>Component</th><th>Status</th><th>Inspection score</th><th>Details</th></tr></thead><tbody>$rows</tbody></table><div class="note"><strong>Buying guide:</strong> 85-100 promising; 70-84 caution; below 70 high risk. Coverage below 90%, battery unknown, warnings or critical failures override the score. Physical checks are required before buying.<br><br><strong>Important:</strong> Supplemental checks (weight 0) are informational and do not affect the health score or diagnostic coverage. Windows events, driver warnings and storage free space may require follow-up. Scores are heuristic inspection ratings, not measured remaining lifespan or probability of reliability. Battery health is capacity retention, not runtime. SMART and firmware may hide disk defects. RAM allocation is not a full memory test. CPU/GPU tests are not long thermal or graphics stress tests. Manually inspect screen, keyboard, hinges, speakers, ports, webcam, charging, battery runtime and seller ownership before buying.</div></body></html>
"@
    $html|Set-Content (Join-Path $OutDir 'report.html') -Encoding UTF8
    Stage 100 'Completed'
}catch{
    $_|Out-String|Set-Content (Join-Path $OutDir 'error.txt') -Encoding UTF8
    Stage 100 'Error: inspection incomplete'
}

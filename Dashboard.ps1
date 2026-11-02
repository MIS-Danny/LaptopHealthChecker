# Laptop Health Checker - Two-Column Dashboard v2.2
# Windows PowerShell 5.1 / Windows 10 and 11
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$bg=[Drawing.Color]::FromArgb(11,18,32)
$panel=[Drawing.Color]::FromArgb(22,34,53)
$panelAlt=[Drawing.Color]::FromArgb(28,43,65)
$fg=[Drawing.Color]::FromArgb(238,244,253)
$muted=[Drawing.Color]::FromArgb(154,170,193)
$blue=[Drawing.Color]::FromArgb(59,130,246)
$green=[Drawing.Color]::FromArgb(74,222,128)
$amber=[Drawing.Color]::FromArgb(251,191,36)
$red=[Drawing.Color]::FromArgb(248,113,113)

$form=New-Object Windows.Forms.Form
$form.Text='Laptop Health Checker | v2.2 Buyer Edition'
$form.Size=New-Object Drawing.Size(950,670)
$form.MinimumSize=New-Object Drawing.Size(850,670)
$form.StartPosition='CenterScreen'
$form.BackColor=$bg
$form.ForeColor=$fg
$form.Font=New-Object Drawing.Font('Segoe UI',10)

function New-Label {
 param([string]$Text,[int]$X,[int]$Y,[int]$W,[int]$H,[int]$Size=10,[bool]$Bold=$false,[Drawing.Color]$Color=$fg,[Windows.Forms.Control]$Parent=$form)
 $l=New-Object Windows.Forms.Label
 $l.Text=$Text;$l.Location=New-Object Drawing.Point($X,$Y);$l.Size=New-Object Drawing.Size($W,$H)
 $style=if($Bold){[Drawing.FontStyle]::Bold}else{[Drawing.FontStyle]::Regular}
 $l.Font=New-Object Drawing.Font('Segoe UI',$Size,$style)
 $l.ForeColor=$Color;$l.BackColor=[Drawing.Color]::Transparent
 $l.AutoEllipsis=$true
 $Parent.Controls.Add($l)
 return $l
}
function New-Card {
 param([int]$X,[int]$Y,[int]$W,[int]$H)
 $p=New-Object Windows.Forms.Panel
 $p.Location=New-Object Drawing.Point($X,$Y);$p.Size=New-Object Drawing.Size($W,$H)
 $p.BackColor=$panel
 $form.Controls.Add($p)
 return $p
}
function New-Button {
 param([string]$Text,[int]$X,[int]$Y,[int]$W,[Drawing.Color]$Color)
 $b=New-Object Windows.Forms.Button
 $b.Text=$Text;$b.Location=New-Object Drawing.Point($X,$Y);$b.Size=New-Object Drawing.Size($W,45)
 $b.FlatStyle='Flat';$b.FlatAppearance.BorderSize=0
 $b.BackColor=$Color;$b.ForeColor=$fg
 $b.Font=New-Object Drawing.Font('Segoe UI',10,[Drawing.FontStyle]::Bold)
 $b.Cursor=[Windows.Forms.Cursors]::Hand
 $form.Controls.Add($b)
 return $b
}

# Two-column buyer dashboard: fixed summary rail on left, scrolling diagnostics on right.
$brand=New-Label 'LAPTOP HEALTH CHECKER' 22 13 630 31 17 $true
$subtitle=New-Label 'PORTABLE HARDWARE INSPECTION  /  WINDOWS 10 & 11' 23 45 650 18 8 $false $muted
$version=New-Label 'V2.2  |  BUYER EDITION' 675 23 235 22 8 $true $muted
$version.TextAlign='MiddleRight'

$scoreCard=New-Card 22 83 276 100
$scoreTitle=New-Label 'HEALTH SCORE' 14 10 240 20 9 $true $muted $scoreCard
$score=New-Label '-- / 100' 14 33 246 51 28 $true $fg $scoreCard

$coverageCard=New-Card 22 193 276 100
$coverageTitle=New-Label 'DIAGNOSTIC COVERAGE' 14 10 240 20 9 $true $muted $coverageCard
$coverage=New-Label '--%' 14 33 246 51 28 $true $fg $coverageCard

$verdictCard=New-Card 22 303 276 102
$verdictTitle=New-Label 'BUYING RECOMMENDATION' 14 9 248 20 9 $true $muted $verdictCard
$verdict=New-Label 'READY TO INSPECT' 14 32 248 59 10 $true $muted $verdictCard
$verdict.TextAlign='MiddleLeft'

$progressLabel=New-Label 'INSPECTION PROGRESS' 22 418 260 18 8 $true $muted
$progress=New-Object Windows.Forms.ProgressBar
$progress.Location=New-Object Drawing.Point(22,440)
$progress.Size=New-Object Drawing.Size(276,8)
$progress.Style='Continuous';$progress.Maximum=100
$form.Controls.Add($progress)
$status=New-Label 'Ready. Choose a test to begin.' 22 456 276 42 9 $false $muted

$quick=New-Button 'QUICK TEST' 22 505 133 $blue
$deep=New-Button 'DEEP TEST' 165 505 133 $panelAlt
$open=New-Button 'HTML REPORT' 22 548 133 $panelAlt
$folder=New-Button 'REPORTS FOLDER' 165 548 133 $panelAlt
foreach($button in @($quick,$deep,$open,$folder)){
 $button.Height=35
 $button.Font=New-Object Drawing.Font('Segoe UI',8,[Drawing.FontStyle]::Bold)
}
$open.Enabled=$false

$resultsTitle=New-Label 'HARDWARE RESULTS' 319 84 320 25 12 $true $fg
$resultsSub=New-Label 'Scroll results; select a row for details' 319 111 555 20 8 $false $muted

$grid=New-Object Windows.Forms.DataGridView
$grid.Location=New-Object Drawing.Point(319,140)
$grid.Size=New-Object Drawing.Size(591,390)
$grid.ReadOnly=$true
$grid.AllowUserToAddRows=$false
$grid.AllowUserToDeleteRows=$false
$grid.AllowUserToResizeRows=$false
$grid.MultiSelect=$false
$grid.RowHeadersVisible=$false
$grid.AutoSizeColumnsMode='Fill'
$grid.BackgroundColor=$panel
$grid.BorderStyle='None'
$grid.GridColor=[Drawing.Color]::FromArgb(47,63,86)
$grid.EnableHeadersVisualStyles=$false
$grid.ColumnHeadersHeight=29
$grid.RowTemplate.Height=27
$grid.SelectionMode='FullRowSelect'
$grid.ColumnHeadersDefaultCellStyle.BackColor=$panelAlt
$grid.ColumnHeadersDefaultCellStyle.ForeColor=$fg
$grid.ColumnHeadersDefaultCellStyle.Font=New-Object Drawing.Font('Segoe UI',9,[Drawing.FontStyle]::Bold)
$grid.DefaultCellStyle.BackColor=$panel
$grid.DefaultCellStyle.ForeColor=$fg
$grid.DefaultCellStyle.SelectionBackColor=[Drawing.Color]::FromArgb(36,66,107)
$grid.DefaultCellStyle.SelectionForeColor=$fg
$grid.DefaultCellStyle.Font=New-Object Drawing.Font('Segoe UI',9)
$grid.DefaultCellStyle.Padding=New-Object Windows.Forms.Padding(6,1,6,1)
$grid.AlternatingRowsDefaultCellStyle.BackColor=$panelAlt
[void]$grid.Columns.Add('Component','COMPONENT')
[void]$grid.Columns.Add('Status','STATUS')
[void]$grid.Columns.Add('Score','SCORE')
[void]$grid.Columns.Add('Details','DETAILS')
$grid.Columns[0].FillWeight=115
$grid.Columns[1].FillWeight=65
$grid.Columns[2].FillWeight=55
$grid.Columns[3].FillWeight=200
$form.Controls.Add($grid)

$detail=New-Label 'Select a component to view complete diagnostic details.' 319 541 590 45 9 $false $muted
$note=New-Label 'Screening estimates only. Verify screen, keyboard, hinges, ports, charging and real battery runtime.' 22 600 890 28 8 $false $muted

# Preserve the left summary rail; give additional width and height to results.
function Update-CompactLayout {
 $w=$form.ClientSize.Width
 $h=$form.ClientSize.Height
 $left=22
 $rail=276
 $right=319
 $rightWidth=[math]::Max(480,$w-$right-22)
 $version.Left=[math]::Max(650,$w-257)
 $grid.Left=$right
 $grid.Width=$rightWidth
 $grid.Height=[math]::Max(275,$h-140-104)
 $resultsTitle.Left=$right
 $resultsSub.Left=$right
 $resultsSub.Width=$rightWidth
 $detail.Left=$right
 $detail.Top=$grid.Bottom+9
 $detail.Width=$rightWidth
 $note.Top=$h-31
 $note.Width=[math]::Max(790,$w-44)
}
$form.Add_Resize({Update-CompactLayout})
Update-CompactLayout

$base=Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'Laptop Health Checker Reports'
New-Item -ItemType Directory -Force -Path $base | Out-Null
$script:runDir=$null
$script:process=$null
$timer=New-Object Windows.Forms.Timer
$timer.Interval=600

$grid.Add_SelectionChanged({
 if($grid.SelectedRows.Count -gt 0){
  $v=$grid.SelectedRows[0].Cells['Details'].Value
  $detail.Text=if($null -ne $v){[string]$v}else{''}
 }
})
$grid.Add_CellFormatting({
 param($sender,$e)
 if($e.RowIndex -lt 0 -or $e.ColumnIndex -ne 1 -or $null -eq $e.Value){return}
 switch([string]$e.Value){
  'PASS' {$e.CellStyle.ForeColor=$green}
  'WARN' {$e.CellStyle.ForeColor=$amber}
  'FAIL' {$e.CellStyle.ForeColor=$red}
  default {$e.CellStyle.ForeColor=$muted}
 }
})

$timer.Add_Tick({
 if(!$script:runDir){return}
 $progressFile=Join-Path $script:runDir 'progress.json'
 if(Test-Path $progressFile){
  try{
   $p=Get-Content -LiteralPath $progressFile -Raw -ErrorAction Stop | ConvertFrom-Json
   $progress.Value=[math]::Min(100,[math]::Max(0,[int]$p.Percent))
   $status.Text=[string]$p.Message
  }catch{}
 }
 if($script:process -and $script:process.HasExited){
  $timer.Stop()
  $quick.Enabled=$true;$deep.Enabled=$true
  $json=Join-Path $script:runDir 'report.json'
  if(Test-Path $json){
   try{
    $r=Get-Content -LiteralPath $json -Raw -ErrorAction Stop | ConvertFrom-Json
    $score.Text=if($r.Overall -ge 0){"$($r.Overall) / 100"}else{'N/A'}
    $score.ForeColor=if($r.Overall -ge 85){$green}elseif($r.Overall -ge 70){$amber}else{$red}
    $coverage.Text="$($r.Coverage)%"
    $verdict.Text=[string]$r.Verdict
    $verdict.ForeColor=if($r.Verdict -match 'DO NOT BUY|HIGH RISK'){$red}elseif($r.Verdict -match 'CAUTION|INCOMPLETE'){$amber}else{$green}
    $grid.Rows.Clear()
    foreach($item in $r.Results){
     $display=if($item.Score -lt 0){'N/A'}else{"$($item.Score)/100"}
     [void]$grid.Rows.Add($item.Name,$item.Status,$display,$item.Detail)
    }
    $grid.ClearSelection()
    $detail.Text='Select a component to view its complete diagnostic details.'
    $open.Enabled=(Test-Path (Join-Path $script:runDir 'report.html'))
    $status.Text="Completed: $($r.Date)  |  $($r.Mode) test"
   }catch{$status.Text='Could not read inspection results: '+$_.Exception.Message}
  }else{
   $status.Text='Inspection failed. Open Reports Folder and check error.txt.'
   $verdict.Text='INSPECTION INCOMPLETE';$verdict.ForeColor=$red
  }
  $script:process.Dispose();$script:process=$null
 }
})

function Start-Test([string]$Mode){
 if($script:process -and !$script:process.HasExited){return}
 $script:runDir=Join-Path $base ((Get-Date -Format 'yyyyMMdd_HHmmss')+'_'+$Mode)
 New-Item -ItemType Directory -Force -Path $script:runDir | Out-Null
 $worker=Join-Path $PSScriptRoot 'Worker.ps1'
 if(!(Test-Path $worker)){
  [Windows.Forms.MessageBox]::Show('Worker.ps1 is missing. Keep Dashboard.ps1 and Worker.ps1 in the same folder.','Missing worker') | Out-Null
  return
 }
 $arguments='-NoProfile -ExecutionPolicy Bypass -File "'+$worker+'" -OutDir "'+$script:runDir+'" -Mode '+$Mode
 $info=New-Object Diagnostics.ProcessStartInfo
 $info.FileName='powershell.exe';$info.Arguments=$arguments
 $info.UseShellExecute=$false;$info.CreateNoWindow=$true
 try{
  $script:process=[Diagnostics.Process]::Start($info)
  $quick.Enabled=$false;$deep.Enabled=$false;$open.Enabled=$false
  $progress.Value=0;$score.Text='-- / 100';$score.ForeColor=$fg
  $coverage.Text='--%';$verdict.Text='RUNNING';$verdict.ForeColor=$blue
  $grid.Rows.Clear();$detail.Text='Inspection in progress...'
  $status.Text='Starting '+$Mode+' test...'
  $timer.Start()
 }catch{
  [Windows.Forms.MessageBox]::Show($_.Exception.Message,'Launch error') | Out-Null
 }
}
$quick.Add_Click({Start-Test 'Quick'})
$deep.Add_Click({Start-Test 'Deep'})
$open.Add_Click({
 $path=Join-Path $script:runDir 'report.html'
 if(Test-Path $path){Start-Process -FilePath $path}
})
$folder.Add_Click({Start-Process -FilePath 'explorer.exe' -ArgumentList ('"'+$base+'"')})
$form.Add_FormClosing({
 if($script:process -and !$script:process.HasExited){
  $answer=[Windows.Forms.MessageBox]::Show('A test is still running. Close the dashboard? The test will continue in the background.','Confirm',[Windows.Forms.MessageBoxButtons]::YesNo)
  if($answer -ne [Windows.Forms.DialogResult]::Yes){$_.Cancel=$true}
 }
})
[void]$form.ShowDialog()

# Laptop Health Checker - Dark Dashboard v2
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
$form.Text='Laptop Health Checker | v2'
$form.Size=New-Object Drawing.Size(1050,760)
$form.MinimumSize=New-Object Drawing.Size(880,680)
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

$brand=New-Label 'LAPTOP HEALTH CHECKER' 30 20 650 38 20 $true
$subtitle=New-Label 'PORTABLE HARDWARE INSPECTION  /  WINDOWS 10 & 11' 32 59 770 24 9 $false $muted
$version=New-Label 'V2  |  SECONDHAND BUYER EDITION' 765 32 245 25 9 $true $muted
$version.TextAlign='MiddleRight'
$version.Anchor='Top,Right'

$scoreCard=New-Card 30 100 470 155
$scoreCard.Anchor='Top,Left'
$scoreTitle=New-Label 'OVERALL INSPECTION SCORE' 22 15 360 23 10 $true $muted $scoreCard
$score=New-Label '-- / 100' 19 41 310 73 37 $true $fg $scoreCard
$scoreHint=New-Label 'Screening score - not a hardware warranty' 23 120 420 22 9 $false $muted $scoreCard
$coverageCard=New-Card 515 100 505 155
$coverageCard.Anchor='Top,Right'
$coverageTitle=New-Label 'DIAGNOSTIC COVERAGE' 22 15 420 24 10 $true $muted $coverageCard
$coverage=New-Label '--%' 20 45 200 63 30 $true $fg $coverageCard
$verdict=New-Label 'READY TO INSPECT' 207 55 270 56 12 $true $muted $coverageCard
$verdict.TextAlign='MiddleRight'
$coverageHint=New-Label 'Unknown items are excluded from the score' 23 120 460 22 9 $false $muted $coverageCard

$progressLabel=New-Label 'INSPECTION PROGRESS' 31 271 300 22 9 $true $muted
$progress=New-Object Windows.Forms.ProgressBar
$progress.Location=New-Object Drawing.Point(30,299)
$progress.Size=New-Object Drawing.Size(990,13)
$progress.Style='Continuous';$progress.Maximum=100
$progress.Anchor='Top,Left,Right'
$form.Controls.Add($progress)
$status=New-Label 'Ready. Choose Quick Test or Deep Test to begin.' 31 320 975 27 10 $false $muted
$status.Anchor='Top,Left,Right'

$quick=New-Button 'QUICK TEST' 30 364 220 $blue
$deep=New-Button 'DEEP TEST' 263 364 220 $panelAlt
$open=New-Button 'OPEN HTML REPORT' 496 364 245 $panelAlt
$folder=New-Button 'REPORTS FOLDER' 754 364 266 $panelAlt
$folder.Anchor='Top,Right'
$open.Enabled=$false

$resultsTitle=New-Label 'HARDWARE RESULTS' 31 428 500 26 12 $true $fg
$resultsSub=New-Label 'Select a row to read its complete details below.' 600 430 420 22 9 $false $muted
$resultsSub.TextAlign='MiddleRight';$resultsSub.Anchor='Top,Right'

$grid=New-Object Windows.Forms.DataGridView
$grid.Location=New-Object Drawing.Point(30,462)
$grid.Size=New-Object Drawing.Size(990,170)
$grid.Anchor='Top,Bottom,Left,Right'
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
$grid.ColumnHeadersHeight=36
$grid.RowTemplate.Height=34
$grid.SelectionMode='FullRowSelect'
$grid.ColumnHeadersDefaultCellStyle.BackColor=$panelAlt
$grid.ColumnHeadersDefaultCellStyle.ForeColor=$fg
$grid.ColumnHeadersDefaultCellStyle.Font=New-Object Drawing.Font('Segoe UI',10,[Drawing.FontStyle]::Bold)
$grid.DefaultCellStyle.BackColor=$panel
$grid.DefaultCellStyle.ForeColor=$fg
$grid.DefaultCellStyle.SelectionBackColor=[Drawing.Color]::FromArgb(36,66,107)
$grid.DefaultCellStyle.SelectionForeColor=$fg
$grid.DefaultCellStyle.Padding=New-Object Windows.Forms.Padding(8,3,8,3)
$grid.AlternatingRowsDefaultCellStyle.BackColor=$panelAlt
[void]$grid.Columns.Add('Component','COMPONENT')
[void]$grid.Columns.Add('Status','STATUS')
[void]$grid.Columns.Add('Score','SCORE')
[void]$grid.Columns.Add('Details','DETAILS')
$grid.Columns[0].FillWeight=100
$grid.Columns[1].FillWeight=64
$grid.Columns[2].FillWeight=56
$grid.Columns[3].FillWeight=255
$form.Controls.Add($grid)

$detail=New-Label 'Component details will appear here after a test.' 32 643 982 36 9 $false $muted
$detail.Anchor='Bottom,Left,Right'
$note=New-Label 'Inspection scores are estimates. Manually verify display, keyboard, hinges, ports, charging and battery runtime before purchase.' 31 688 984 33 9 $false $muted
$note.Anchor='Bottom,Left,Right'

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
    $score.ForeColor=if($r.Overall -ge 80){$green}elseif($r.Overall -ge 60){$amber}else{$red}
    $coverage.Text="$($r.Coverage)%"
    $verdict.Text=[string]$r.Verdict
    $verdict.ForeColor=if($r.Verdict -match 'AVOID'){$red}elseif($r.Verdict -match 'INSPECT|INSUFFICIENT'){$amber}else{$green}
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

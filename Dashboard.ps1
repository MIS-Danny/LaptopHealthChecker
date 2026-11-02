Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
$form=New-Object Windows.Forms.Form
$form.Text='Laptop Health Checker | Portable'
$form.Size=New-Object Drawing.Size(850,670)
$form.StartPosition='CenterScreen';$form.BackColor=[Drawing.Color]::FromArgb(244,247,251);$form.Font=New-Object Drawing.Font('Segoe UI',10)
function Label([string]$text,[int]$x,[int]$y,[int]$w,[int]$h,[int]$size=10,[bool]$bold=$false){$l=New-Object Windows.Forms.Label;$l.Text=$text;$l.Location=New-Object Drawing.Point($x,$y);$l.Size=New-Object Drawing.Size($w,$h);$l.Font=New-Object Drawing.Font('Segoe UI',$size,$(if($bold){[Drawing.FontStyle]::Bold}else{[Drawing.FontStyle]::Regular}));$form.Controls.Add($l);return $l}
$title=Label 'LAPTOP HEALTH CHECKER' 25 18 600 40 19 $true
$sub=Label 'Portable secondhand laptop inspection  |  Windows 10 / 11' 27 58 740 27 10 $false
$score=Label '-- / 100' 27 100 280 65 32 $true
$verdict=Label 'Ready to inspect' 330 113 470 42 17 $true
$coverage=Label 'Coverage: --' 29 169 600 28
$progress=New-Object Windows.Forms.ProgressBar;$progress.Location=New-Object Drawing.Point(28,204);$progress.Size=New-Object Drawing.Size(775,20);$form.Controls.Add($progress)
$status=Label 'Choose Quick Test or Deep Test.' 28 232 760 30
$quick=New-Object Windows.Forms.Button;$quick.Text='Quick Test';$quick.Location=New-Object Drawing.Point(28,273);$quick.Size=New-Object Drawing.Size(170,42);$form.Controls.Add($quick)
$deep=New-Object Windows.Forms.Button;$deep.Text='Deep Test';$deep.Location=New-Object Drawing.Point(211,273);$deep.Size=New-Object Drawing.Size(170,42);$form.Controls.Add($deep)
$open=New-Object Windows.Forms.Button;$open.Text='Open HTML Report';$open.Location=New-Object Drawing.Point(394,273);$open.Size=New-Object Drawing.Size(194,42);$open.Enabled=$false;$form.Controls.Add($open)
$folder=New-Object Windows.Forms.Button;$folder.Text='Open Reports Folder';$folder.Location=New-Object Drawing.Point(600,273);$folder.Size=New-Object Drawing.Size(203,42);$form.Controls.Add($folder)
$grid=New-Object Windows.Forms.DataGridView;$grid.Location=New-Object Drawing.Point(28,333);$grid.Size=New-Object Drawing.Size(775,230);$grid.ReadOnly=$true;$grid.AllowUserToAddRows=$false;$grid.AllowUserToDeleteRows=$false;$grid.RowHeadersVisible=$false;$grid.AutoSizeColumnsMode='Fill';$grid.BackgroundColor=[Drawing.Color]::White;$grid.SelectionMode='FullRowSelect';$grid.Columns.Add('Component','Component')|Out-Null;$grid.Columns.Add('Status','Status')|Out-Null;$grid.Columns.Add('Score','Score')|Out-Null;$grid.Columns.Add('Details','Details')|Out-Null;$grid.Columns[0].FillWeight=85;$grid.Columns[1].FillWeight=55;$grid.Columns[2].FillWeight=45;$grid.Columns[3].FillWeight=210;$form.Controls.Add($grid)
$note=Label 'Important: Scores are screening estimates, not guaranteed hardware health. Test screen, keyboard, ports and battery runtime manually.' 28 577 780 47 9
$base=Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'Laptop Health Checker Reports'
New-Item -ItemType Directory -Force -Path $base|Out-Null
$script:runDir=$null;$script:process=$null
$timer=New-Object Windows.Forms.Timer;$timer.Interval=600
$timer.Add_Tick({
 if(!$script:runDir){return}
 $progressFile=Join-Path $script:runDir 'progress.json'
 if(Test-Path $progressFile){try{$p=Get-Content $progressFile -Raw|ConvertFrom-Json;$progress.Value=[math]::Min(100,[math]::Max(0,[int]$p.Percent));$status.Text=$p.Message}catch{}}
 if($script:process -and $script:process.HasExited){
  $timer.Stop();$quick.Enabled=$true;$deep.Enabled=$true
  $json=Join-Path $script:runDir 'report.json'
  if(Test-Path $json){
   try{$r=Get-Content $json -Raw|ConvertFrom-Json;$score.Text=if($r.Overall -ge 0){"$($r.Overall) / 100"}else{'N/A'};$verdict.Text=$r.Verdict;$coverage.Text="Diagnostic coverage: $($r.Coverage)%";$grid.Rows.Clear();foreach($item in $r.Results){$display=if($item.Score -lt 0){'N/A'}else{"$($item.Score)/100"};[void]$grid.Rows.Add($item.Name,$item.Status,$display,$item.Detail)};$open.Enabled=$true;$status.Text="Completed: $($r.Date)  |  $($r.Mode)"}catch{$status.Text='Could not read inspection results: '+$_.Exception.Message}
  }else{$status.Text='Inspection failed. Check error.txt in report folder.'}
  $script:process.Dispose();$script:process=$null
 }
})
function Start-Test([string]$mode){
 if($script:process -and !$script:process.HasExited){return}
 $script:runDir=Join-Path $base ((Get-Date -Format 'yyyyMMdd_HHmmss')+'_'+$mode)
 New-Item -ItemType Directory -Force -Path $script:runDir|Out-Null
 $worker=Join-Path $PSScriptRoot 'Worker.ps1'
 $args='-NoProfile -ExecutionPolicy Bypass -File "'+$worker+'" -OutDir "'+$script:runDir+'" -Mode '+$mode
 $info=New-Object Diagnostics.ProcessStartInfo;$info.FileName='powershell.exe';$info.Arguments=$args;$info.UseShellExecute=$false;$info.CreateNoWindow=$true
 try{$script:process=[Diagnostics.Process]::Start($info);$quick.Enabled=$false;$deep.Enabled=$false;$open.Enabled=$false;$progress.Value=0;$score.Text='-- / 100';$verdict.Text='Inspection running';$grid.Rows.Clear();$status.Text='Starting '+$mode+' test...';$timer.Start()}catch{[Windows.Forms.MessageBox]::Show($_.Exception.Message,'Launch error')|Out-Null}
}
$quick.Add_Click({Start-Test 'Quick'})
$deep.Add_Click({Start-Test 'Deep'})
$open.Add_Click({$path=Join-Path $script:runDir 'report.html';if(Test-Path $path){Start-Process $path}})
$folder.Add_Click({Start-Process explorer.exe -ArgumentList ('"'+$base+'"')})
$form.Add_FormClosing({if($script:process -and !$script:process.HasExited){$answer=[Windows.Forms.MessageBox]::Show('A test is still running. Close the dashboard? The test will continue in the background.','Confirm',[Windows.Forms.MessageBoxButtons]::YesNo);if($answer -ne 'Yes'){$_.Cancel=$true}}})
[void]$form.ShowDialog()

#requires -Version 5.1
# Synthetic report parity check; no database connection is used.
$ErrorActionPreference='Stop'
$source=Join-Path (Split-Path $PSScriptRoot -Parent) 'csdirector-database-project-insights.ps1'
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile($source,[ref]$tokens,[ref]$errors)
if($errors.Count){throw ($errors | Out-String)}
$functions=@($ast.FindAll({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst]},$true))
foreach($fn in $functions){. ([scriptblock]::Create($fn.Extent.Text))}

$Client=$false;$Database='synthetic';$IncludeTrussJson=$true;$TimingLog=$false
$LegacyCutoffVersion='2024.7.2.7';$LegacyState='Before 2024.r7';$cut=ParseVer $LegacyCutoffVersion
$Groups=@('Identity','Version');$Map=[ordered]@{Quantity='ComponentQuantity'}
$MapMeta=[ordered]@{Quantity=[ordered]@{group='Identity';basis='observed mapping';policy='int'}}
$Rules=[ordered]@{
 'R-PROJ'=@('Project not in DB','Synthetic rule');'R-DB'=@('Truss missing from DB','Synthetic rule')
 'R-TDL'=@('Truss file missing','Synthetic rule');'R-FCI'=@('Check-in times differ','Synthetic rule')
 'R-CKSUM'=@('Truss file checksum differs','Synthetic rule');'R-VER'=@('Truss version differs','Synthetic rule')
 'R-FLD'=@('Truss value differs','Synthetic rule');'R-FILE'=@('Truss missing from archive','Synthetic rule')
}
$Badge=@{'Before 2024.r7'='b-pre';'DB only'='b-pre';'Not in DB'='b-warn'}
$Css='';$ProjectTemplate='<html>@@DATA@@</html>';$modeLabel='TEST'
$script:DbCountCache=[Collections.Generic.Dictionary[string,long]]::new([StringComparer]::OrdinalIgnoreCase)
$root=Join-Path ([IO.Path]::GetTempPath()) ('insights-parallel-test-'+[guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $root)
Add-Type -AssemblyName System.IO.Compression.FileSystem
$archive=Join-Path $root 'P3_20251112000000.csdproj'
$zip=[IO.Compression.ZipFile]::Open($archive,[IO.Compression.ZipArchiveMode]::Create)
try{
 $entry=$zip.CreateEntry('Trusses/T1.tdlTruss')
 $writer=[IO.StreamWriter]::new($entry.Open())
 try{$writer.Write('synthetic truss')}finally{$writer.Dispose()}
}finally{$zip.Dispose()}

function SetSyntheticBatch{
 $script:ProjectBatch=[Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase)
 $script:TrussBatch=[Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase)
 $script:ProjectBatch['P0']=$null
 foreach($number in @('P1','P2')){
  $script:ProjectBatch[$number]=$null
  $names=[Collections.Generic.Dictionary[string,object]]::new()
  $rows=[Collections.Generic.List[object]]::new()
  $rows.Add(@{Name='T1';AnalysisVersion=$(if($number -eq 'P1'){'2023.1'}else{'2025.1'});CompatibilityVersion='1';TrussFileCheckSum='01-02';ComponentQuantity=2})
  $names['T1']=$rows
  $script:TrussBatch[$number]=$names
 }
 $script:ProjectBatch['P3']=$null
 $archiveRows=[Collections.Generic.List[object]]::new()
 $checksum=[BitConverter]::ToString([Security.Cryptography.MD5]::Create().ComputeHash([Text.Encoding]::UTF8.GetBytes('synthetic truss')))
 $archiveRows.Add(@{Name='T1';AnalysisVersion='2025.1';CompatibilityVersion='1';TrussFileCheckSum=$checksum;ComponentQuantity=2})
 $archiveNames=[Collections.Generic.Dictionary[string,object]]::new()
 $archiveNames['T1']=$archiveRows
 $script:TrussBatch['P3']=$archiveNames
}

try{
 $outDir=Join-Path $root 'sequential';[void](New-Item -ItemType Directory -Path (Join-Path $outDir 'projects') -Force)
 $used=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
 $results=[Collections.Generic.List[object]]::new()
 SetSyntheticBatch
 foreach($number in @('P0','P1','P2')){WriteProject (AnalyzeDbOnly $number $null $cut) $number}
 WriteProject (AnalyzeProject 'P3' $archive $cut) 'P3'
 $expected=@($results | ForEach-Object {$_.project+':'+$_.state} | Sort-Object)

 $outDir=Join-Path $root 'parallel';[void](New-Item -ItemType Directory -Path (Join-Path $outDir 'projects') -Force)
 $used=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
 $results=[Collections.Generic.List[object]]::new()
 SetSyntheticBatch
 $state=[Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
 foreach($fn in $functions){
  $body=$fn.Body.Extent.Text
  $parameters=if($fn.Parameters.Count){'param('+(@($fn.Parameters | ForEach-Object {$_.Extent.Text}) -join ',')+")`n"}else{''}
  $state.Commands.Add([Management.Automation.Runspaces.SessionStateFunctionEntry]::new($fn.Name,$parameters+$body.Substring(1,$body.Length-2)))
 }
 foreach($name in @('Client','Database','IncludeTrussJson','LegacyCutoffVersion','LegacyState','Badge','Groups','Map','MapMeta','Rules','Css','ProjectTemplate','modeLabel','outDir')){
  $state.Variables.Add([Management.Automation.Runspaces.SessionStateVariableEntry]::new($name,(Get-Variable $name -ValueOnly),'test'))
 }
 $state.Variables.Add([Management.Automation.Runspaces.SessionStateVariableEntry]::new('TimingLog',$false,'test'))
 $state.Variables.Add([Management.Automation.Runspaces.SessionStateVariableEntry]::new('ParallelWorker',$true,'test'))
 $parallelPool=[runspacefactory]::CreateRunspacePool(1,2,$state,$Host);$parallelPool.Open()
 $pending=[Collections.Generic.List[object]]::new()
 $TimingLog=$true;$script:TimingStart=[Diagnostics.Stopwatch]::StartNew();$script:TimingPath=Join-Path $root 'timing.csv'
 $startOutput=@(foreach($number in @('P0','P1','P2')){StartParallelProject $number $null $true $null 1 'db_only'};StartParallelProject 'P3' $archive $false $null 3 'archive')
 if($startOutput.Count){throw ('Worker start leaked '+$startOutput.Count+' output objects')}
 $finishOutput=@(while($pending.Count){FinishParallelProject})
 if($finishOutput.Count){throw ('Worker completion leaked '+$finishOutput.Count+' output objects')}
 if(@(Get-Content $script:TimingPath).Count -ne 12){throw 'Missing parallel timing rows'}
 if((@($results | ForEach-Object {$_.project+':'+$_.state} | Sort-Object) -join '|') -ne ($expected -join '|')){throw 'State mismatch'}
 foreach($number in @('P0','P1','P2','P3')){
  $files=@('summary.json','findings.json','project.html')
  if($number -ne 'P0'){$files+='trusses\T1.json'}
  foreach($file in $files){
   $a=[IO.File]::ReadAllBytes((Join-Path $root ('sequential\projects\'+$number+'\'+$file)))
   $b=[IO.File]::ReadAllBytes((Join-Path $root ('parallel\projects\'+$number+'\'+$file)))
   if([Convert]::ToBase64String($a) -ne [Convert]::ToBase64String($b)){throw ('Output mismatch: '+$number+' '+$file)}
  }
 }
 Write-Host 'Parallel DB-only and archive output matches sequential output.'
}finally{
 if($parallelPool){$parallelPool.Dispose()}
 $full=[IO.Path]::GetFullPath($root)
 $temp=[IO.Path]::GetFullPath([IO.Path]::GetTempPath())
 if($full.StartsWith($temp,[StringComparison]::OrdinalIgnoreCase)){Remove-Item -LiteralPath $full -Recurse -Force}
}

#requires -Version 5.1
<#
.SYNOPSIS
Read-only comparison of CSDirector database trusses with the TrussStudio files inside .csdproj archives. Produces a summary index, one page per project and JSON for each project and truss.
.EXAMPLE
.\csdirector-database-project-insights.ps1 -SqlUser '<sql-user>' -Server '<server-or-instance>' -ProjectsRoot 'C:\path\to\projects'
.EXAMPLE
.\csdirector-database-project-insights.ps1 -ProjectId 'PROJECT-001','PROJECT-002'
.EXAMPLE
.\csdirector-database-project-insights.ps1 -Limit 5
.EXAMPLE
.\csdirector-database-project-insights.ps1 -IncludeDbOnly -TimingLog -Limit 100
.NOTES
With -SqlUser the password is read only from the CSD_SQL_PASSWORD environment variable. Archives and SQL records are never changed.
Projects whose trusses are all older than -LegacyCutoffVersion (pre 2024.r7) are DB-only and are not counted as inconsistent.
A project whose DB check-in id is older than the archive timestamp is reported as Stale DB rather than Inconsistent.
#>
[CmdletBinding()]
param(
 [switch]$Client,
 [string]$ProjectsRoot=$(if($Client){'C:\SST\Client\Projects'}else{'C:\SST\Server\Projects'}),
 [string]$Server='.\SST',
 [string]$SqlUser,
 [string]$Database=$(if($Client){'CSDatabaseClient'}else{'CSDatabaseServer'}),
 [string]$OutputDirectory=(Join-Path (Get-Location).Path 'output-insights'),
 [string[]]$ProjectId,
 [ValidatePattern('^\d+(\.\d+){0,3}$')][string]$LegacyCutoffVersion='2024.7.2.7',
 [ValidateRange(0,2147483647)][int]$Limit=0,
 [switch]$IncludeDbOnly,
 [switch]$IncludeTrussJson,
 [switch]$TimingLog
)
$ErrorActionPreference='Stop'
Write-Host ''
Write-Host 'CSDirector: Database and Project file Insights' -ForegroundColor Cyan
Write-Host ('Compares DB trusses with TrussStudio files in each .csdproj using '+$(if($SqlUser){'SQL authentication ('+$SqlUser+').'}else{'Windows authentication.'}))
Write-Host ('Target: '+$(if($Client){'Client'}else{'Server'})+' database and projects.')
Write-Host 'Archives and database records are read only.'
Write-Host 'Press Enter to accept each default shown in brackets. Command-line values skip their prompts.'
Write-Host ''
foreach($inputName in @('ProjectsRoot','Server','Database','OutputDirectory')) {
 if(-not $PSBoundParameters.ContainsKey($inputName)) {
  $defaultValue=Get-Variable -Name $inputName -ValueOnly
  $label=switch($inputName){
   'ProjectsRoot' {'Projects folder'}
   'Server' {'SQL server / instance'}
   'Database' {'Database name'}
   'OutputDirectory' {'Output folder base (timestamp appended)'}
  }
  $answer=Read-Host ($label+' ['+$defaultValue+']')
  if(-not [string]::IsNullOrWhiteSpace($answer)){Set-Variable -Name $inputName -Value $answer.Trim()}
 }
}
Add-Type -AssemblyName System.IO.Compression.FileSystem
if($Server -ieq 'SST'){$Server='.\SST'}
$script:TrussColumns=$null
$script:DbCountCache=[Collections.Generic.Dictionary[string,long]]::new([StringComparer]::OrdinalIgnoreCase)
$Groups=@('Identity','Version','Geometry','Pitch','Flags')
$Map=[ordered]@{
 'Quantity'='ComponentQuantity';'NumPlies'='TrussPlyCount';'Thickness'='TrussThicknessInches';'OverallTrussHeight'='OverallTrussHeightInches';'LeftOverhang'='LeftOverhangInches';'RightOverhang'='RightOverhangInches';'LeftCantilever'='LeftCantileverInches';'RightCantilever'='RightCantileverInches';'LeftHeelHeight'='LeftHeelHeightInches';'RightHeelHeight'='RightHeelHeightInches';'PitchLeftTopOver12'='LeftTopPitchOver12';'PitchRightTopOver12'='RightTopPitchOver12';'PitchLeftBottomOver12'='LeftBottomPitchOver12';'PitchRightBottomOver12'='RightBottomPitchOver12';'IsAttic'='IsAttic';'IsGable'='IsGable';'IsGirder'='IsGirder';'IsFlipped'='IsFlipped';'OCSpacing'='OCSpacing'
 'Height'='TrussHeightInches';'LeftHeelToPeak'='LeftHeelToPeakInches';'RightHeelToPeak'='RightHeelToPeakInches';'LeftTopChordHeelToPeak'='LeftTopChordHeelToPeakInches';'RightTopChordHeelToPeak'='RightTopChordHeelToPeakInches';'LeftStub'='LeftStubInches';'RightStub'='RightStubInches';'TCPitchCount'='TopChordPitchCount';'BCPitchCount'='BottomChordPitchCount';'Camber'='Camber';'CamberDistance'='CamberDistanceInches';'CamberNote'='CamberNote';'BearingPadNotes'='BearingPadNotes';'LeftOverhangRake'='LeftOverhangRake';'RightOverhangRake'='RightOverhangRake';'SinglePlyOverhang'='SinglePlyOverhang';'SinglePlyFill'='SinglePlyFill';'FillPly'='FillPly';'OverhangPly'='OverhangPly';'TCRakeFirst'='TCRakeFirst';'TCRakeLast'='TCRakeLast';'CompatibilityVersion'='CompatibilityVersion';'TrussMatchCode'='TrussMatchCode'
 'Span'='TrussSpanInches';'ProdSpan'='FullSpanInches';'IsStructural'='IsStructuralGable';'Weight'='WeightPounds';'PlyToPlyNote'='PlyToPly';'IsSymmetric'='IsSymmetrical'
}
$MapMeta=[ordered]@{
 'Quantity'=[ordered]@{group='Identity';policy='int';basis='observed mapping'}
 'NumPlies'=[ordered]@{group='Geometry';policy='int';basis='observed mapping'}
 'Thickness'=[ordered]@{group='Geometry';policy='numeric';basis='observed mapping'}
 'OverallTrussHeight'=[ordered]@{group='Geometry';policy='numeric';basis='observed mapping'}
 'LeftOverhang'=[ordered]@{group='Geometry';policy='numeric';basis='observed mapping'}
 'RightOverhang'=[ordered]@{group='Geometry';policy='numeric';basis='observed mapping'}
 'LeftCantilever'=[ordered]@{group='Geometry';policy='numeric';basis='observed mapping'}
 'RightCantilever'=[ordered]@{group='Geometry';policy='numeric';basis='observed mapping'}
 'LeftHeelHeight'=[ordered]@{group='Geometry';policy='numeric';basis='observed mapping'}
 'RightHeelHeight'=[ordered]@{group='Geometry';policy='numeric';basis='observed mapping'}
 'PitchLeftTopOver12'=[ordered]@{group='Pitch';policy='numeric';basis='observed mapping'}
 'PitchRightTopOver12'=[ordered]@{group='Pitch';policy='numeric';basis='observed mapping'}
 'PitchLeftBottomOver12'=[ordered]@{group='Pitch';policy='numeric';basis='observed mapping'}
 'PitchRightBottomOver12'=[ordered]@{group='Pitch';policy='numeric';basis='observed mapping'}
 'IsAttic'=[ordered]@{group='Flags';policy='bool';basis='observed mapping'}
 'IsGable'=[ordered]@{group='Flags';policy='bool';basis='observed mapping'}
 'IsGirder'=[ordered]@{group='Flags';policy='bool';basis='observed mapping'}
 'IsFlipped'=[ordered]@{group='Flags';policy='bool';basis='observed mapping'}
 'OCSpacing'=[ordered]@{group='Geometry';policy='numeric';basis='observed mapping'}
 'Height'=[ordered]@{group='Geometry';policy='numeric';basis='observed mapping'}
 'LeftHeelToPeak'=[ordered]@{group='Geometry';policy='numeric';basis='observed mapping'}
 'RightHeelToPeak'=[ordered]@{group='Geometry';policy='numeric';basis='observed mapping'}
 'LeftTopChordHeelToPeak'=[ordered]@{group='Geometry';policy='numeric';basis='observed mapping'}
 'RightTopChordHeelToPeak'=[ordered]@{group='Geometry';policy='numeric';basis='observed mapping'}
 'LeftStub'=[ordered]@{group='Geometry';policy='numeric';basis='observed mapping'}
 'RightStub'=[ordered]@{group='Geometry';policy='numeric';basis='observed mapping'}
 'TCPitchCount'=[ordered]@{group='Pitch';policy='int';basis='observed mapping'}
 'BCPitchCount'=[ordered]@{group='Pitch';policy='int';basis='observed mapping'}
 'Camber'=[ordered]@{group='Geometry';policy='numeric';basis='observed mapping'}
 'CamberDistance'=[ordered]@{group='Geometry';policy='numeric';basis='observed mapping'}
 'CamberNote'=[ordered]@{group='Geometry';policy='text';basis='observed mapping'}
 'BearingPadNotes'=[ordered]@{group='Geometry';policy='text';basis='observed mapping'}
 'LeftOverhangRake'=[ordered]@{group='Geometry';policy='numeric';basis='observed mapping'}
 'RightOverhangRake'=[ordered]@{group='Geometry';policy='numeric';basis='observed mapping'}
 'SinglePlyOverhang'=[ordered]@{group='Flags';policy='bool';basis='observed mapping'}
 'SinglePlyFill'=[ordered]@{group='Flags';policy='bool';basis='observed mapping'}
 'FillPly'=[ordered]@{group='Geometry';policy='int';basis='observed mapping'}
 'OverhangPly'=[ordered]@{group='Geometry';policy='int';basis='observed mapping'}
 'TCRakeFirst'=[ordered]@{group='Geometry';policy='numeric';basis='observed mapping'}
 'TCRakeLast'=[ordered]@{group='Geometry';policy='numeric';basis='observed mapping'}
 'CompatibilityVersion'=[ordered]@{group='Version';policy='text';basis='observed mapping'}
 'TrussMatchCode'=[ordered]@{group='Geometry';policy='text';basis='observed mapping'}
 'Span'=[ordered]@{group='Geometry';policy='numeric';basis='documented persistence mapping'}
 'ProdSpan'=[ordered]@{group='Geometry';policy='numeric';basis='documented persistence mapping'}
 'IsStructural'=[ordered]@{group='Flags';policy='bool';basis='documented persistence mapping'}
 'Weight'=[ordered]@{group='Geometry';policy='numeric';basis='documented persistence mapping'}
 'PlyToPlyNote'=[ordered]@{group='Geometry';policy='text';basis='documented persistence mapping'}
 'IsSymmetric'=[ordered]@{group='Flags';policy='bool';basis='documented persistence mapping'}
}
$Rules=[ordered]@{
 'R-PROJ'=@('Project not in DB','Archive exists but no Project row with this number in the DB.')
 'R-FCI'=@('Check-in times differ','DB and archive represent different check-ins; this alone does not establish a truss data mismatch.')
 'R-FILE'=@('Truss missing from archive','DB truss at a current (r7+) version has no Trusses/<name>.tdlTruss in the archive; the .tdlTruss controls membership and JSON is never membership. Before 2024.r7 DB-only trusses are expected and excluded.')
 'R-DB'=@('Truss missing from DB','Trusses/<name>.tdlTruss exists in the archive but no ComponentHeader row with that name; the .tdlTruss controls membership and JSON is never membership.')
 'R-DUP'=@('Duplicate truss in DB','More than one DB row for the same truss name.')
 'R-CKSUM'=@('Truss file checksum differs','MD5 of Trusses/<name>.tdlTruss should equal ComponentHeader.TrussFileCheckSum.')
 'R-TDL'=@('Truss file missing','No .tdlTruss to checksum.')
 'R-VER'=@('Truss version differs','Truss JSON Version should equal ComponentTruss.AnalysisVersion. Reported as stale when the DB is frozen before 2024.r7 and the file is current.')
 'R-FLD'=@('Truss value differs','Mapped truss JSON field should equal its ComponentHeader/Component/ComponentTruss column (tolerance 0.001).')
}
$StaleRules=@('R-CKSUM','R-DB','R-FLD','R-VER','R-FILE')
$Badge=@{'Inconsistent'='b-bad';'Consistent'='b-ok';'Before 2024.r7'='b-pre';'Stale DB'='b-warn';'Not in DB'='b-warn';'DB only'='b-pre'}

function ParseVer($s){
 if($null -eq $s){return $null}
 $parts=([string]$s).Trim().Split('.')
 if($parts.Count -gt 4){$parts=$parts[0..3]}
 [decimal]$key=0;$k=0
 foreach($x in $parts){
  $i=0
  if(-not [int]::TryParse($x,[ref]$i)){return $null}
  $key=$key*1000000+$i;$k++
 }
 while($k -lt 4){$key=$key*1000000;$k++}
 return $key
}
function VerLess($a,$b){$a -lt $b}
function Esc($s){[Net.WebUtility]::HtmlEncode([string]$s)}
function SafeName($n){$s=[regex]::Replace($n,'[^A-Za-z0-9._-]+','_');if($s.Length -gt 80){$s=$s.Substring(0,80)};$s}
function WriteText($path,$text){[IO.File]::WriteAllText($path,$text,[Text.UTF8Encoding]::new($false))}
function ToJson($o,[int]$depth=8,[switch]$compress){
 if($compress){ConvertTo-Json -InputObject $o -Depth $depth -Compress}else{ConvertTo-Json -InputObject $o -Depth $depth}
}
function LogTiming([string]$phase,[double]$durationMs=0,[long]$rows=0,[long]$trusses=0,[double]$serializeMs=0,[double]$writeMs=0,[long]$htmlBytes=0){
 if(-not $TimingLog){return}
 $inv=[Globalization.CultureInfo]::InvariantCulture
 $proc=[Diagnostics.Process]::GetCurrentProcess()
 try{
  $proc.Refresh()
  $values=@(
   [DateTime]::UtcNow.ToString('o',$inv),
   $script:TimingStart.Elapsed.TotalMilliseconds.ToString('F1',$inv),
   $phase,
   [string]$script:DiagnosticProjectIndex,
   [string]$script:DiagnosticKind,
   [string]$rows,
   [string]$trusses,
   $durationMs.ToString('F1',$inv),
   $serializeMs.ToString('F1',$inv),
   $writeMs.ToString('F1',$inv),
   [string]$htmlBytes,
   ([GC]::GetTotalMemory($false)/1MB).ToString('F1',$inv),
   ($proc.WorkingSet64/1MB).ToString('F1',$inv),
   ($proc.PrivateMemorySize64/1MB).ToString('F1',$inv),
   $proc.TotalProcessorTime.TotalMilliseconds.ToString('F1',$inv),
   [string][GC]::CollectionCount(0),
   [string][GC]::CollectionCount(1),
   [string][GC]::CollectionCount(2)
  )
  [IO.File]::AppendAllText($script:TimingPath,($values -join ',')+"`r`n",[Text.UTF8Encoding]::new($false))
 }finally{$proc.Dispose()}
}
function Cells($v){$c=[ordered]@{};foreach($g in $Groups){$c[$g]=$v};$c}
function IdentityOnly($state){$c=Cells 'na';$c['Identity']=$state;$c}
function Q($sql,$params,[string]$phase='sql_query'){
 $cmd=$script:Conn.CreateCommand();$cmd.CommandText=$sql;$cmd.CommandTimeout=120
 foreach($k in $params.Keys){[void]$cmd.Parameters.AddWithValue('@'+$k,$params[$k])}
 $dt=[Data.DataTable]::new()
 $adapter=[Data.SqlClient.SqlDataAdapter]::new($cmd)
 try{
  if($TimingLog){LogTiming ($phase+'_start');$timer=[Diagnostics.Stopwatch]::StartNew()}
  [void]$adapter.Fill($dt)
  if($TimingLog){$timer.Stop();LogTiming $phase $timer.Elapsed.TotalMilliseconds $dt.Rows.Count}
 }finally{$adapter.Dispose();$cmd.Dispose()}
 return ,$dt
}
function AddF($row,$All,$stale,$rule,$field,$dbv,$filev,$msg,$grp,$sev='error'){
 if($StaleRules -contains $rule -and ($stale -or $sev -eq 'warn')){$sev='stale'}
 $item=[ordered]@{rule=$rule;truss=$row.name;field=$field;db=$dbv;file=$filev;sev=$sev;msg=$msg;group=$grp}
 $row.findings.Add($item);$All.Add($item)
 $nv=if($sev -eq 'stale'){'stale'}else{'bad'}
 $c=$row.cells[$grp]
 if($c -ne 'bad' -and $c -ne 'pre' -and ($nv -eq 'bad' -or $c -eq 'ok' -or $c -eq 'na')){$row.cells[$grp]=$nv}
}
function CmpPair($a,$b,$policy){
 switch($policy){
  'bool'{
   $an=NormalizeBool $a;$bn=NormalizeBool $b
   if($null -ne $an -and $null -ne $bn){return ($an -eq $bn)}
  }
  'int'{
   $ai=0;$bi=0
   if([int]::TryParse([string]$a,[ref]$ai) -and [int]::TryParse([string]$b,[ref]$bi)){return ($ai -eq $bi)}
  }
  'numeric'{
   $x=0.0;$y=0.0
   if([double]::TryParse(([string]$a),[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$x) -and [double]::TryParse(([string]$b),[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$y)){return ([math]::Abs($x-$y) -le 0.001)}
  }
 }
 [string]::Equals([string]$a,[string]$b,[StringComparison]::Ordinal)
}
function NormalizeBool($v){
 if($v -is [bool]){return $v}
 if($v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [decimal]){
  $d=[double]$v
  if($d -eq 0){return $false}
  if($d -eq 1){return $true}
  return $null
 }
 $s=[string]$v
 if($s -eq '0'){return $false}
 if($s -eq '1'){return $true}
 $null
}
function SideInfo($o,$name){
 if($null -eq $o){return [ordered]@{exists=$false;isNull=$false;value=$null;whole=$true}}
 $p=$o.PSObject.Properties[$name]
 if($null -eq $p){return [ordered]@{exists=$false;isNull=$false;value=$null;whole=$false}}
 [ordered]@{exists=$true;isNull=($null -eq $p.Value);value=$p.Value;whole=$false}
}
function DbSide($r,$col){
 if($null -eq $r){return [ordered]@{exists=$false;isNull=$false;value=$null;whole=$true}}
 if(-not $r.Contains($col)){return [ordered]@{exists=$false;isNull=$false;value=$null;whole=$false}}
 $b=$r[$col]
 [ordered]@{exists=$true;isNull=($null -eq $b);value=$b;whole=$false}
}
function FieldStatus($f,$d,$candidate){
 if($candidate){return 'unresolved'}
 $fe=($f.exists -and -not $f.isNull)
 $de=($d.exists -and -not $d.isNull)
 if($fe -and $de){return 'compared'}
 if($fe){
  if($d.whole){return 'missing side / not compared'}
  if($d.exists){return 'DB NULL / not compared'}
  return 'missing DB column / not compared'
 }
 if($de){
  if($f.whole){return 'missing side / not compared'}
  if($f.isNull){return 'file NULL / not compared'}
  return 'missing field / not compared'
 }
 if($d.whole -or $f.whole){return 'missing side / not compared'}
 if($d.isNull -and $f.isNull){return 'both null / not compared'}
 if($d.exists){return 'missing field / not compared'}
 'missing DB column / not compared'
}
function CoverageOf($fields){
 $c=[ordered]@{compared=0;different=0;notCompared=0}
 foreach($f in $fields){
  if($f.status -eq 'match'){$c.compared++}
  elseif($f.status -eq 'differ'){$c.compared++;$c.different++}
  else{$c.notCompared++}
 }
 $c
}
function ReadEntryBytes($entry){
 if($entry -is [IO.FileInfo]){return ,[IO.File]::ReadAllBytes($entry.FullName)}
 $s=$entry.Open();$ms=[IO.MemoryStream]::new()
 try{$s.CopyTo($ms)}finally{$s.Dispose()}
 return ,$ms.ToArray()
}
function ReadEntryJson($entry){
 if($entry -is [IO.FileInfo]){return ([IO.File]::ReadAllText($entry.FullName,[Text.Encoding]::UTF8) | ConvertFrom-Json)}
 $s=$entry.Open();$r=[IO.StreamReader]::new($s,[Text.Encoding]::UTF8,$true)
 try{$t=$r.ReadToEnd()}finally{$r.Dispose()}
 $t | ConvertFrom-Json
}
function DbBatchEnd($items,[int]$start){
 $last=[Math]::Min($start+49,$items.Count-1)
 $params=@{};$names=[Collections.Generic.List[string]]::new()
 for($i=$start;$i -le $last;$i++){
  $number=[string]$items[$i][0]
  if($script:DbCountCache.ContainsKey($number)){continue}
  $script:DbCountCache[$number]=0L
  $key='p'+$names.Count;$params[$key]=$number;$names.Add('@'+$key)
 }
 if($names.Count -gt 0){
  $dt=Q ('SELECT p.ProjectNumber,COUNT_BIG(*) AS TrussCount FROM dbo.Project p JOIN dbo.ComponentHeader h ON h.ProjectEventKey=p.EventKey LEFT JOIN dbo.Component c ON c.ComponentKey=h.ComponentHeaderKey LEFT JOIN dbo.ComponentTruss t ON t.ComponentKey=c.ComponentKey WHERE p.ProjectNumber IN ('+($names -join ',')+') GROUP BY p.ProjectNumber') $params 'sql_count'
  foreach($r in $dt.Rows){$script:DbCountCache[[string]$r['ProjectNumber']]=[long]$r['TrussCount']}
 }
 $total=0L;$end=$start
 for($i=$start;$i -le $last;$i++){
  $count=$script:DbCountCache[[string]$items[$i][0]]
  if($i -gt $start -and $total+$count -gt 10000){break}
  $total+=$count;$end=$i
 }
 $end
}
function ReleaseProject($number){
 [void]$script:TrussBatch.Remove($number)
 [void]$script:ProjectBatch.Remove($number)
 [void]$script:DbCountCache.Remove($number)
}
function LoadDbBatch([string[]]$numbers){
 if($numbers.Count -gt 50){throw 'DB batch exceeds 50 projects.'}
 $script:ProjectBatch=[Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase)
 $script:TrussBatch=[Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase)
 if(-not $numbers.Count){return}
 $params=@{};$placeholders=for($i=0;$i -lt $numbers.Count;$i++){$params['p'+$i]=$numbers[$i];'@p'+$i}
 $filter=$placeholders -join ','
 $projects=Q ('SELECT p.ProjectNumber,p.FilesCheckInId FROM dbo.Project p WHERE p.ProjectNumber IN ('+$filter+')') $params 'sql_projects'
 foreach($r in $projects.Rows){
  $number=[string]$r['ProjectNumber']
  $script:ProjectBatch[$number]=$(if($r['FilesCheckInId'] -is [DBNull]){$null}else{[string]$r['FilesCheckInId']})
 }
 if($null -eq $script:TrussColumns){
  $schema=Q "SELECT name FROM sys.columns WHERE object_id=OBJECT_ID(N'dbo.ComponentTruss')" @{} 'sql_schema'
  if(-not $schema.Rows.Count){throw 'Cannot read dbo.ComponentTruss column metadata; comparison cannot safely continue.'}
  $script:TrussColumns=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
  foreach($column in $schema.Rows){[void]$script:TrussColumns.Add([string]$column['name'])}
 }
 $columns=@('AnalysisVersion','CompatibilityVersion')+@($Map.Values | Where-Object {$_ -notin @('ComponentQuantity','WeightPounds','IsSymmetrical')})
 $select=(@($columns | Sort-Object -Unique | Where-Object {$script:TrussColumns.Contains($_)} | ForEach-Object {',t.['+$_.Replace(']',']]')+']'}) -join '')
 $dt=Q ('SELECT p.ProjectNumber,h.Name,h.TrussFileCheckSum,h.WeightPounds,h.IsSymmetrical,c.ComponentQuantity'+$select+' FROM dbo.Project p JOIN dbo.ComponentHeader h ON h.ProjectEventKey=p.EventKey LEFT JOIN dbo.Component c ON c.ComponentKey=h.ComponentHeaderKey LEFT JOIN dbo.ComponentTruss t ON t.ComponentKey=c.ComponentKey WHERE p.ProjectNumber IN ('+$filter+') ORDER BY p.ProjectNumber,h.Name') $params 'sql_trusses'
 if($TimingLog){LogTiming 'db_materialize_start' 0 $dt.Rows.Count;$materializeTimer=[Diagnostics.Stopwatch]::StartNew()}
 foreach($r in $dt.Rows){
  $number=[string]$r['ProjectNumber']
  if(-not $script:TrussBatch.ContainsKey($number)){$script:TrussBatch[$number]=[Collections.Generic.Dictionary[string,object]]::new()}
  $d=$script:TrussBatch[$number];$o=@{}
  foreach($col in $dt.Columns){
   if($col.ColumnName -eq 'ProjectNumber'){continue}
   $v=$r[$col.ColumnName]
   if($v -is [DBNull]){$o[$col.ColumnName]=$null}
   elseif($v -is [byte[]]){if($col.ColumnName -eq 'TrussFileCheckSum'){$o[$col.ColumnName]=[BitConverter]::ToString($v)}}
   elseif($v -is [guid] -or $v -is [datetime]){$o[$col.ColumnName]=[string]$v}
   else{$o[$col.ColumnName]=$v}
  }
  $n=[string]$o['Name']
  if(-not $d.ContainsKey($n)){$d[$n]=[Collections.Generic.List[object]]::new()}
  $d[$n].Add($o)
 }
 if($TimingLog){$materializeTimer.Stop();LogTiming 'db_materialize' $materializeTimer.Elapsed.TotalMilliseconds $dt.Rows.Count}
}
function DbTrusses($number){
 if($script:TrussBatch.ContainsKey($number)){return ,$script:TrussBatch[$number]}
 return ,([Collections.Generic.Dictionary[string,object]]::new())
}
function FillFields($row,$j,$r,$h,$F,$stale,$cut){
 $jv=$null;$jc=$null;$dv=$null;$dc=$null
 if($j){$pv=$j.PSObject.Properties['Version'];if($pv){$jv=$pv.Value};$pv=$j.PSObject.Properties['CompatibilityVersion'];if($pv){$jc=$pv.Value}}
 if($r){$dv=$r['AnalysisVersion'];$dc=$r['CompatibilityVersion']}
 $row.version=[ordered]@{file=$jv;db=$dv;compat_file=$jc;compat_db=$dc}
 $vs=SideInfo $j 'Version'
 $ds=DbSide $r 'AnalysisVersion'
 $vst=FieldStatus $vs $ds $false
 $vsame=$null
 if($vst -eq 'compared'){
  $vsame=([string]::Equals([string]$jv,[string]$dv,[StringComparison]::Ordinal))
  if(-not $vsame){
   $fvp=ParseVer $jv;$dvp=ParseVer $dv
   if($null -ne $dvp -and (VerLess $dvp $cut) -and -not ($null -ne $fvp -and (VerLess $fvp $cut))){
    AddF $row $F $stale 'R-VER' 'Version vs AnalysisVersion' $dv $jv 'DB frozen at before 2024.r7 version; file is current (re-opened in Studio, DB not refreshed)' 'Version' 'warn'
   }else{
    AddF $row $F $stale 'R-VER' 'Version vs AnalysisVersion' $dv $jv 'File and DB version differ' 'Version'
   }
   $vst='differ'
  }else{$vst='match'}
 }
 $row.fields.Add([ordered]@{field='Version';group='Version';column='AnalysisVersion';db=$dv;file=$jv;match=$vsame;status=$vst})
 if($vst -eq 'match' -and $row.cells['Version'] -eq 'na'){$row.cells['Version']='ok'}
 $dk=$null;if($r){$dk=$r['TrussFileCheckSum']}
 $cm=$null
 if($null -ne $dk -and $h){$cm=(([string]$dk).ToUpperInvariant() -eq $h)}
 $cs=[ordered]@{exists=($null -ne $h);isNull=($null -eq $h);value=$h;whole=$false}
 $cst=FieldStatus $cs (DbSide $r 'TrussFileCheckSum') $false
 if($cst -eq 'compared'){$cst=$(if($cm){'match'}else{'differ'})}
 $row.fields.Add([ordered]@{field='TdlTrussMD5';group='Identity';column='TrussFileCheckSum';db=$dk;file=$h;match=$cm;status=$cst})
 if($cst -eq 'match' -and $row.cells['Identity'] -eq 'na'){$row.cells['Identity']='ok'}
 if($null -ne $r -and $cm -eq $false){AddF $row $F $stale 'R-CKSUM' 'TrussFileCheckSum' ([string]$dk) $h '.tdlTruss MD5 differs from DB checksum' 'Identity'}
 foreach($fld in $Map.Keys){
  $meta=$MapMeta[$fld]
  $col=$Map[$fld]
  $grp=$meta.group
  $candidate=($meta.basis -eq 'candidate')
  $a=SideInfo $j $fld
  $b=DbSide $r $col
  $st=FieldStatus $a $b $candidate
  $eq=$null
  if($st -eq 'compared'){
   $eq=CmpPair $a.value $b.value $meta.policy
   $st=$(if($eq){'match'}else{'differ'})
   if($row.cells[$grp] -eq 'na'){$row.cells[$grp]='ok'}
   if(-not $eq){AddF $row $F $stale 'R-FLD' $fld $b.value $a.value ($col+' differs') $grp}
  }
  $row.fields.Add([ordered]@{field=$fld;group=$grp;column=$col;db=$b.value;file=$a.value;match=$eq;status=$st;basis=$meta.basis;policy=$meta.policy})
 }
 $row.coverage=CoverageOf $row.fields
}
function AnalyzeDbOnly($number,$ci){
 $out=[ordered]@{project=$number;archive=$null;archiveTs=$null;dbFilesCheckInId=$ci;notInDb=$false;dbOnly=$true;trusses=[Collections.Generic.List[object]]::new();findings=[Collections.Generic.List[object]]::new()}
 $dbt=DbTrusses $number
 $out.counts=[ordered]@{trusses=$dbt.Count;findings=0;bad=0;stale=0;pre=0}
 $out.staleProject=$false
 return ,$out
}
function FillDbOnlyFields($row,$r){
 # With no archive side, mapped values cannot be compared. Keep every field and
 # status in the report without running the per-field comparison functions.
 $dv=$r['AnalysisVersion'];$dc=$r['CompatibilityVersion']
 $row.version=[ordered]@{file=$null;db=$dv;compat_file=$null;compat_db=$dc}
 $row.fields.Add([ordered]@{field='Version';group='Version';column='AnalysisVersion';db=$dv;file=$null;match=$null;status='missing side / not compared'})
 $dk=$r['TrussFileCheckSum']
 $cs=[ordered]@{exists=$false;isNull=$true;value=$null;whole=$false}
 $cst=FieldStatus $cs (DbSide $r 'TrussFileCheckSum') $false
 $row.fields.Add([ordered]@{field='TdlTrussMD5';group='Identity';column='TrussFileCheckSum';db=$dk;file=$null;match=$null;status=$cst})
 foreach($fld in $Map.Keys){
  $meta=$MapMeta[$fld];$col=$Map[$fld]
  $dbv=$null;if($r.Contains($col)){$dbv=$r[$col]}
  $status=if($meta.basis -eq 'candidate'){'unresolved'}else{'missing side / not compared'}
  $row.fields.Add([ordered]@{field=$fld;group=$meta.group;column=$col;db=$dbv;file=$null;match=$null;status=$status;basis=$meta.basis;policy=$meta.policy})
 }
 $row.coverage=[ordered]@{compared=0;different=0;notCompared=$row.fields.Count}
}
function CompactDbOnlyRow($name,$r){
 $values=[Collections.Generic.List[object]]::new()
 foreach($fld in $Map.Keys){
  $col=$Map[$fld]
  if($r.Contains($col)){$values.Add($r[$col])}else{$values.Add($null)}
 }
 $checksumStatus=if(-not $r.Contains('TrussFileCheckSum')){'missing DB column / not compared'}elseif($null -eq $r['TrussFileCheckSum']){'both null / not compared'}else{'file NULL / not compared'}
 return ,([object[]]@($name,$r['AnalysisVersion'],$r['CompatibilityVersion'],$r['TrussFileCheckSum'],$checksumStatus,$values.ToArray()))
}
function AnalyzeProject($number,$archivePath,$cut){
 $isDir=Test-Path -LiteralPath $archivePath -PathType Container
 $tsName=[IO.Path]::GetFileName($archivePath)
 if($isDir){
  $tsName=''
  $zs=@(Get-ChildItem -LiteralPath $archivePath -Filter '*.csdproj' -File | Where-Object {$_.Name -notmatch '^Attachments?_'} | Sort-Object Name)
  if($zs.Count -gt 0){$tsName=$zs[$zs.Count-1].Name}
 }
 $m=[regex]::Match($tsName,'_(\d{14})')
 $ts=if($m.Success){$m.Groups[1].Value}else{$null}
 $hasProject=$script:ProjectBatch.ContainsKey($number)
 $dbCheckIn=$null
 if($hasProject){$dbCheckIn=$script:ProjectBatch[$number]}
 $dbt=[Collections.Generic.Dictionary[string,object]]::new()
 if($hasProject){$dbt=DbTrusses $number}
 $out=[ordered]@{project=$number;archive=$(if($isDir){[IO.Path]::GetFileName($archivePath)+' (expanded folder)'}else{[IO.Path]::GetFileName($archivePath)});archiveTs=$ts;dbFilesCheckInId=$dbCheckIn;notInDb=(-not $hasProject);trusses=[Collections.Generic.List[object]]::new();findings=[Collections.Generic.List[object]]::new()}
 $F=$out.findings
 if(-not $hasProject){
  $F.Add([ordered]@{rule='R-PROJ';truss=$null;field='Project';db=$null;file=$number;sev='error';msg='Project not in DB';group='Project'})
 }elseif(-not $Client -and $ts -and $dbCheckIn -ne $ts){
  $F.Add([ordered]@{rule='R-FCI';truss=$null;field='FilesCheckInId';db=$dbCheckIn;file=$ts;sev='warn';msg='DB check-in id differs from archive timestamp';group='Project'})
 }
 $stale=[bool](-not $Client -and $hasProject -and $ts -and $null -ne $dbCheckIn -and [string]::CompareOrdinal($dbCheckIn,$ts) -lt 0)
 $zip=$null
 if(-not $isDir){$zip=[IO.Compression.ZipFile]::OpenRead($archivePath)}
 try{
  $files=[Collections.Generic.Dictionary[string,object]]::new()
  $tdl=[Collections.Generic.Dictionary[string,object]]::new()
  if($isDir){
   $gen=Join-Path $archivePath 'Trusses\generated'
   if(Test-Path -LiteralPath $gen -PathType Container){foreach($e in (Get-ChildItem -LiteralPath $gen -Filter '*.json' -File)){$files[$e.BaseName]=$e}}
   $tdir=Join-Path $archivePath 'Trusses'
   if(Test-Path -LiteralPath $tdir -PathType Container){foreach($e in (Get-ChildItem -LiteralPath $tdir -Filter '*.tdlTruss' -File)){$tdl[$e.BaseName]=$e}}
  }else{
   foreach($e in $zip.Entries){
    $mm=[regex]::Match($e.FullName,'^Trusses/generated/(.+)\.json$')
    if($mm.Success){$files[$mm.Groups[1].Value]=$e;continue}
    $mm=[regex]::Match($e.FullName,'^Trusses/(.+)\.tdlTruss$')
    if($mm.Success){$tdl[$mm.Groups[1].Value]=$e}
   }
  }
  $names=[Collections.Generic.List[string]]::new()
  foreach($k in $tdl.Keys){$names.Add($k)}
  foreach($k in $dbt.Keys){if(-not $tdl.ContainsKey($k)){$names.Add($k)}}
  $names.Sort([StringComparer]::Ordinal)
  foreach($name in $names){
   $row=[ordered]@{name=$name;cells=(Cells 'na');findings=[Collections.Generic.List[object]]::new();fields=[Collections.Generic.List[object]]::new();version=$null;pre=$false}
   $rows=$null
   if($dbt.ContainsKey($name)){$rows=$dbt[$name]}
   $r=$null;if($rows){$r=$rows[0]}
   $j=$null;$h=$null
   if($tdl.ContainsKey($name) -and $files.ContainsKey($name)){$j=ReadEntryJson $files[$name]}
   if($tdl.ContainsKey($name)){
    $md5=[Security.Cryptography.MD5]::Create()
    try{$h=[BitConverter]::ToString($md5.ComputeHash((ReadEntryBytes $tdl[$name])))}finally{$md5.Dispose()}
   }
   if(-not $tdl.ContainsKey($name)){
    $av=$null
    if($r){$av=ParseVer $r['AnalysisVersion']}
    if($r -and $null -ne $av -and (VerLess $av $cut)){
     $row.cells=Cells 'pre';$row.pre=$true
    }else{
     AddF $row $F $stale 'R-FILE' ('Trusses/'+$name+'.tdlTruss') 'present' 'missing' 'Truss in DB but no .tdlTruss in archive' 'Identity'
     $row.cells=IdentityOnly $row.cells['Identity']
    }
    FillFields $row $null $r $h $F $stale $cut
    $out.trusses.Add($row);continue
   }
   if(-not $r){
    AddF $row $F $stale 'R-DB' 'ComponentHeader.Name' 'missing' $name 'Truss in file but not in DB' 'Identity'
    $row.cells=IdentityOnly $row.cells['Identity']
    FillFields $row $j $null $h $F $stale $cut
    $out.trusses.Add($row);continue
   }
   if($rows.Count -gt 1){AddF $row $F $stale 'R-DUP' 'ComponentHeader.Name' $rows.Count 1 'Duplicate truss name rows in DB' 'Identity'}
   if($null -eq $j){
    $av=ParseVer $r['AnalysisVersion']
    if($null -ne $av -and (VerLess $av $cut)){$row.cells=Cells 'pre';$row.pre=$true}
   }
   FillFields $row $j $r $h $F $stale $cut
   $out.trusses.Add($row)
  }
 }finally{if($zip){$zip.Dispose()}}
 $bad=0;$st=0;$pre=0
 foreach($t in $out.trusses){
  if($t.pre){$pre++}
  if($t.findings.Count -gt 0){
   $allStale=$true
   foreach($tf in $t.findings){if($tf.sev -ne 'stale'){$allStale=$false}}
   if($allStale){$st++}else{$bad++}
  }
 }
 $out.counts=[ordered]@{trusses=$out.trusses.Count;findings=$out.findings.Count;bad=$bad;stale=$st;pre=$pre}
 $out.staleProject=$stale
 return ,$out
}
function StateOf($o){
 if($o.dbOnly){return 'DB only'}
 if($o.notInDb){return 'Not in DB'}
 foreach($f in $o.findings){if($f.sev -eq 'error'){return 'Inconsistent'}}
 if($o.counts.trusses -gt 0 -and $o.counts.pre -eq $o.counts.trusses){return 'Before 2024.r7'}
 if($o.findings.Count -gt 0){return 'Stale DB'}
 'Consistent'
}
function RuleName($f){
 if($f.rule -ne 'R-FCI'){return $Rules[$f.rule][0]}
 if([string]::IsNullOrWhiteSpace([string]$f.db)){return 'DB check-in is unknown'}
 'Check-in times differ'
}
function CauseDefinitions{
 $causes=[ordered]@{
  'Check-in times differ'='Project.FilesCheckInId and the archive filename timestamp identify different check-ins. This alone does not prove a truss data mismatch. If the archive is newer, applicable truss differences are marked stale.'
  'DB check-in is unknown'='Project.FilesCheckInId is missing; the relative age of the DB and archive cannot be determined.'
 }
 foreach($key in $Rules.Keys){if($key -notin @('R-FCI','R-TDL')){$causes[$Rules[$key][0]]=$Rules[$key][1]}}
 return ,$causes
}
function CauseLegend{
 $items=foreach($entry in (CauseDefinitions).GetEnumerator()){'<div><b>'+(Esc $entry.Key)+'</b> &mdash; '+(Esc $entry.Value)+'</div>'}
 '<footer class="cause-legend"><h3>How to read this report</h3><p>&#10003; Compared values match &middot; &#8800; Real mismatch &middot; &#8635; Stale DB (archive newer) &middot; &#9208; Before 2024.r7 (expected DB-only data) &middot; &#8709; Not comparable.</p><p>Consistent means the compared values agree, not that every field was verified. Missing JSON means values are not compared; only .tdlTruss files establish file-side truss existence. Check-in dates come from identifiers, not file modification times; no timezone is assumed. Client mode does not compare check-ins.</p><details><summary>Cause definitions</summary>'+($items -join '')+'</details></footer>'
}
function VerdictOf($o,$st){
 $n=$o.counts.trusses
 switch($st){
  'Consistent'{'No differences in compared values; see coverage for values not compared.'}
  'Not in DB'{'Archive exists, but the project is not in the DB'}
  'DB only'{'In the DB, but no archive folder on disk. DB values only; nothing to compare.'}
  'Before 2024.r7'{'Frozen before 2024.r7 DB state; no Studio files to compare.'}
  'Stale DB'{if($o.staleProject){"Check-in times differ. $($o.counts.stale) of $n trusses have stale differences."}else{'Check-in times differ; see legend.'}}
  default{"$($o.counts.bad) of $n trusses have real mismatches."}
 }
}

$Css=@'
body{font:14px/1.4 Segoe UI,Arial,sans-serif;margin:0;background:#f4f6f8;color:#1b2733}
header{background:#12304a;color:#fff;padding:14px 24px}header h1{margin:0;font-size:20px}header small{opacity:.75}
main{padding:16px 24px}a{color:#0b5cad}
.badge{display:inline-block;padding:2px 9px;border-radius:11px;font-size:12px;font-weight:600;margin-right:4px;white-space:nowrap}
.b-ok{background:#d6f1de;color:#0d5c28}.b-bad{background:#fbd9d6;color:#9b1c12}.b-pre{background:#e3e6ea;color:#4a5560}.b-warn{background:#fdebc8;color:#7a4b00}
.cards{display:flex;gap:12px;margin-bottom:14px;flex-wrap:wrap}.card{background:#fff;border-radius:6px;padding:10px 16px;box-shadow:0 1px 2px #0002;min-width:120px}.card b{font-size:22px;display:block}
.tabs{display:flex;gap:10px;margin-bottom:14px;flex-wrap:wrap}.tab{cursor:pointer;background:#fff;border-radius:6px;padding:10px 16px;min-width:150px;box-shadow:0 1px 2px #0002;border:2px solid transparent}.tab:hover{border-color:#c9d2db}.tab.on{border-color:#2f6fab}.tab b{font-size:24px;display:block}.tab span{font-size:12px;color:#5a6672}
input[type=text]{padding:6px 9px;border:1px solid #b8c2cc;border-radius:4px;width:260px}
table{border-collapse:collapse;background:#fff;width:100%;box-shadow:0 1px 2px #0002}th,td{padding:6px 10px;border-bottom:1px solid #e3e8ed;text-align:left;vertical-align:top}th{background:#eef1f4;position:sticky;top:0}
.chip{display:inline-block;background:#eef1f4;border-radius:4px;padding:1px 6px;margin:1px;font-size:12px;cursor:pointer;border:1px solid transparent}.chip.on{border-color:#0b5cad;background:#dcebfa}
.cell{width:90px;text-align:center;white-space:nowrap;cursor:pointer;font-weight:700}.c-ok{background:#d6f1de;color:#0d5c28}.c-bad{background:#f4b9b3;color:#8a160c}.c-pre{background:#d5d9de;color:#4a5560}.c-stale{background:#fdebc8;color:#7a4b00}.c-na{background:#f3f3f3;color:#aaa}
.cause-legend{margin-top:24px;padding:12px 16px;border-top:1px solid #c9d2db;color:#4a5560;font-size:12px}.cause-legend h3{margin:0;font-size:14px}.cause-legend p{margin:6px 0}.cause-legend summary{cursor:pointer;font-weight:600}.cause-legend details div{margin:6px 0}
.legend span{margin-right:14px;font-size:12px}.sw{display:inline-block;width:12px;height:12px;vertical-align:-1px;margin-right:4px;border-radius:2px}
.sbs td{font-family:Consolas,monospace;font-size:12px;word-break:break-all}.sbs .d{background:#fbe3e0}.sbs .f{background:#e3f1e6}
.dd td{font-weight:600}.note{background:#fff8e1;border-left:4px solid #e0a800;padding:6px 10px;margin:8px 0;font-size:13px}
'@

$ProjectTemplate=@'
<!doctype html><meta charset=utf-8><title>@@TITLE@@</title><style>@@CSS@@</style>
<header><h1><a href="../../index.html" style="color:#9cd0ff">All projects</a> / @@TITLE@@ <span class="badge @@BADGE@@">@@STATE@@</span><span class="badge b-pre" style="float:right">@@MODE@@</span></h1>
<small>@@META@@</small></header>
<main><div class="legend"><span><i class="sw c-ok"></i>matches</span><span><i class="sw c-bad"></i>real mismatch</span><span><i class="sw c-stale"></i>stale DB (file newer)</span><span><i class="sw c-pre"></i>before 2024.r7 DB-only (expected)</span><span><i class="sw c-na"></i>not comparable</span></div>
<p>Filter by issue: @@CHIPS@@</p>
<p><input type=text id=q placeholder="filter truss name"> <label><input type=checkbox id=onlyBad checked> only trusses with findings</label> <span id=cnt></span></p>
<table style="table-layout:fixed"><thead><tr><th>Truss</th>@@HEADERS@@<th class=cell style="width:220px">Version (file / DB)</th></tr></thead><tbody id=tb></tbody></table>@@LEGEND@@</main>
<script>
const G=@@GROUPS@@,D=@@DATA@@,DB=@@DBONLY@@,M=@@DBMETA@@,NA=G.map(()=> 'na'),CL={ok:'c-ok',bad:'c-bad',stale:'c-stale',pre:'c-pre',na:'c-na'},SY={ok:'&#10003;',bad:'&#8800;',stale:'&#8635;',pre:'&#9208;',na:'&#8709;'},open_=new Set(),all_=new Set();
let rules=new Set();
function esc(s){return String(s==null?'':s).replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]))}
function fv(x){return x==null?'-':(typeof x==='object'?esc(JSON.stringify(x)):esc(x))}
function dbRow(r){const F=[{field:'Version',group:'Version',column:'AnalysisVersion',db:r[1],file:null,match:null,status:'missing side / not compared'},{field:'TdlTrussMD5',group:'Identity',column:'TrussFileCheckSum',db:r[3],file:null,match:null,status:r[4]}];M.forEach((m,i)=>F.push({field:m[0],group:m[1],column:m[2],db:r[5][i],file:null,match:null,status:m[3]==='candidate'?'unresolved':'missing side / not compared',basis:m[3],policy:m[4]}));return [r[0],NA,[],{file:null,db:r[1],compat_file:null,compat_db:r[2]},F,{compared:0,different:0,notCompared:M.length+2}]}
function det(i){const r=DB?dbRow(D[i]):D[i],F=r[4]||[],CV=r[5]||{};let h='<tr><td colspan='+(G.length+1)+' style="background:#fafbfc">';
if(CV.compared!==undefined)h+='<p>Coverage: '+CV.compared+' compared, '+CV.different+' different, '+CV.notCompared+' not compared</p>';
if(r[2].length)h+='<p class=note>'+r[2].map(f=>'<b>'+f[6]+'</b>: '+f[0]+' - '+esc(f[5])+(f[4]=='stale'?' (stale DB)':'')).filter((v,k,a)=>a.indexOf(v)==k).join('<br>')+'</p>';
const bad=F.filter(f=>f.match===false),rest=F.filter(f=>f.match!==false);
const row=(f)=>'<tr class="'+(f.match===false?'dd':'')+'"><td>'+esc(f.group)+'</td><td>'+esc(f.field)+'</td><td>'+esc(f.column)+'</td><td class="'+(f.match===false?'d':f.match?'f':'')+'">'+fv(f.db)+'</td><td class="'+(f.match===false?'d':f.match?'f':'')+'">'+fv(f.file)+'</td><td>'+esc(f.status||(f.match===false?'differs':f.match?'match':'n/a'))+(f.basis==='candidate'?'<br><small>candidate mapping not confirmed; values retained, never flagged</small>':'')+'</td></tr>';
const extra=r[2].filter(f=>f[0]!='Truss value differs'&&f[0]!='Truss version differs');
const xr=extra.map(f=>'<tr class=dd><td>'+f[6]+'</td><td>'+esc(f[1])+'</td><td>-</td><td class=d>'+fv(f[2])+'</td><td class=d>'+fv(f[3])+'</td><td>'+(f[4]=='stale'?'stale':'differs')+'</td></tr>').join('');
if(F.length||xr){h+='<table class=sbs><tr><th>Group</th><th>JSON field</th><th>DB column</th><th>DB</th><th>File (archive)</th><th>Status</th></tr>'+xr+bad.map(row).join('')+(all_.has(i)||!bad.length?rest.map(row).join(''):'')+'</table>';
if(rest.length&&bad.length)h+='<a href="#" onclick="all_.has('+i+')?all_.delete('+i+'):all_.add('+i+');draw();return false">'+(all_.has(i)?'hide':'show')+' '+rest.length+' matching/other fields</a>';}
else h+='<p>No values recorded.</p>';
return h+'</td></tr>'}
function draw(){const q=document.getElementById('q').value.toLowerCase(),ob=document.getElementById('onlyBad').checked;let n=0,h='';
D.forEach((r,i)=>{const findings=DB?[]:r[2],cells=DB?NA:r[1],v=DB?{file:null,db:r[1]}:r[3];if(ob&&!findings.length)return;if(q&&r[0].toLowerCase().indexOf(q)<0)return;if(rules.size&&!findings.some(f=>rules.has(f[0])))return;n++;if(n>400)return;
h+='<tr style="cursor:pointer" onclick="tg('+i+')"><td>'+(open_.has(i)?'&#9662; ':'&#9656; ')+esc(r[0])+'</td>'+cells.map((c,k)=>G[k]=='Version'?'':'<td class="cell '+CL[c]+'">'+SY[c]+'</td>').join('')+'<td class="cell '+CL[cells[1]]+'">'+SY[cells[1]]+(v.file||v.db?' '+esc(v.file||'-')+' / '+esc(v.db||'-'):'')+'</td></tr>'+(open_.has(i)?det(i):'')});
document.getElementById('tb').innerHTML=h;document.getElementById('cnt').textContent=n+' trusses'+(n>400?' (first 400 shown)':'')}
function tg(i){open_.has(i)?open_.delete(i):open_.add(i);draw()}
document.getElementById('q').oninput=draw;document.getElementById('onlyBad').onchange=draw;
document.querySelectorAll('.chip').forEach(c=>c.onclick=()=>{const r=c.dataset.r;rules.has(r)?rules.delete(r):rules.add(r);c.classList.toggle('on');draw()});
if(DB||!D.some(r=>r[2].length))document.getElementById('onlyBad').checked=false;
draw();
</script>
'@

$IndexTemplate=@'
<!doctype html><meta charset=utf-8><title>CSDirector / TrussStudio Insights</title><style>@@CSS@@</style>
<header><h1>CSDirector / TrussStudio Insights <span class="badge b-warn">@@COUNT@@ projects</span><span class="badge b-pre" style="float:right">@@MODE@@</span></h1><small>Summary only. Open a project for the truss-by-truss field comparison. Before 2024.r7 DB-only projects are not counted as inconsistent; Stale DB means the file is newer than the DB, not a data fault. Generated @@WHEN@@ from @@SOURCE@@. Aggregate @@AGGC@@. Contract: <a href="contract.json">contract.json</a>.</small></header>
<main><div class="tabs">@@TABS@@</div>
<p><input type=text id=q placeholder="filter project number"> Issue: @@CHIPS@@</p>
<table><thead><tr><th>Project</th><th>State</th><th>Verdict</th><th>Top causes</th><th>Trusses</th></tr></thead><tbody id=tb>@@ROWS@@</tbody></table>
@@LEGEND@@</main>
<script>let t='Inconsistent';const cs=new Set();function f(){const q=document.getElementById('q').value.toLowerCase();document.querySelectorAll('#tb tr').forEach(r=>{const rc=(r.dataset.c||'').split('|');r.style.display=(r.dataset.s==t&&r.textContent.toLowerCase().indexOf(q)>=0&&(!cs.size||[...cs].some(c=>rc.includes(c))))?'':'none'})}
document.querySelectorAll('.tab').forEach(e=>e.onclick=()=>{t=e.dataset.t;document.querySelectorAll('.tab').forEach(x=>x.classList.toggle('on',x==e));f()});document.querySelectorAll('.chip').forEach(e=>e.onclick=()=>{const c=e.dataset.c;cs.has(c)?cs.delete(c):cs.add(c);e.classList.toggle('on');f()});document.getElementById('q').oninput=f;</script>
'@

if($Server -ieq 'SST'){$Server='.\SST'}
$cut=ParseVer $LegacyCutoffVersion
$modeLabel=$(if($Client){'CLIENT'}else{'SERVER'})+' | '+$Database
$root=(Get-Item -LiteralPath $ProjectsRoot).FullName.TrimEnd('\')
$stamp=(Get-Date).ToString('yyyyMMdd-HHmmss')
$outDir=$OutputDirectory.TrimEnd('\')+'-'+$stamp
[void](New-Item -ItemType Directory -Path (Join-Path $outDir 'projects') -Force)
if($TimingLog){
 $script:TimingStart=[Diagnostics.Stopwatch]::StartNew()
 $script:TimingPath=Join-Path $outDir 'timing.csv'
 $script:DiagnosticProjectIndex=0
 $script:DiagnosticKind='setup'
 WriteText $script:TimingPath "utc,elapsed_ms,phase,project_index,kind,rows,trusses,duration_ms,serialize_ms,write_ms,html_bytes,heap_mb,working_mb,private_mb,cpu_ms,gc0,gc1,gc2`r`n"
 LogTiming 'start'
 Write-Host ('Timing log: '+$script:TimingPath)
}
WriteText (Join-Path $outDir 'rules.json') (ToJson ([ordered]@{}+@{}) 3)
$ruleJson=[ordered]@{}
foreach($k in $Rules.Keys){$ruleJson[$k]=[ordered]@{name=$Rules[$k][0];rule=$Rules[$k][1]}}
WriteText (Join-Path $outDir 'rules.json') (ToJson $ruleJson 3)

$builder=[Data.SqlClient.SqlConnectionStringBuilder]::new()
$builder['Data Source']=$Server;$builder['Initial Catalog']=$Database;$builder['Application Name']='csdirector-database-project-insights'
if($SqlUser){
 if(-not $env:CSD_SQL_PASSWORD){throw 'Set the CSD_SQL_PASSWORD environment variable when using -SqlUser.'}
 $builder['Integrated Security']=$false;$builder['User ID']=$SqlUser;$builder['Password']=$env:CSD_SQL_PASSWORD
}else{$builder['Integrated Security']=$true}
$script:Conn=[Data.SqlClient.SqlConnection]::new($builder.ConnectionString)
if($TimingLog){LogTiming 'sql_connect_start';$connectTimer=[Diagnostics.Stopwatch]::StartNew()}
$script:Conn.Open()
if($TimingLog){$connectTimer.Stop();LogTiming 'sql_connect' $connectTimer.Elapsed.TotalMilliseconds}

$scanTimer=[Diagnostics.Stopwatch]::StartNew()
$work=@()
foreach($d in (Get-ChildItem -LiteralPath $root -Directory | Sort-Object Name)){
 if($d.Name -ieq 'DeletedProjects'){continue}
 if($ProjectId -and ($ProjectId -notcontains $d.Name)){continue}
 if($Client -and (Test-Path -LiteralPath (Join-Path $d.FullName 'Trusses') -PathType Container)){$work+=,@($d.Name,$d.FullName);continue}
 $arc=@(Get-ChildItem -LiteralPath $d.FullName -Filter '*.csdproj' -File | Where-Object {$_.Name -notmatch '^Attachments?_'} | Sort-Object Name)
 if($arc.Count -eq 0){continue}
 $work+=,@($d.Name,$arc[$arc.Count-1].FullName)
}
if($Limit -gt 0){$work=@($work | Select-Object -First $Limit)}
$scanTimer.Stop()
LogTiming 'archive_discovery' $scanTimer.Elapsed.TotalMilliseconds $work.Count
Write-Host ('Projects with archives: '+$work.Count)

function WriteProject($o,$number){
 if($TimingLog){
  LogTiming 'report_start' 0 0 $o.counts.trusses
  $reportTimer=[Diagnostics.Stopwatch]::StartNew()
  $script:ReportSerializeTicks=0L
  $script:ReportWriteTicks=0L
 }
 $st=StateOf $o
 $o.state=$st
 $folder=SafeName $number;$n=2
 while(-not $used.Add($folder)){$folder=(SafeName $number)+'_'+$n;$n++}
 $o.folder=$folder
 $dir=Join-Path $outDir ('projects\'+$folder)
 [void](New-Item -ItemType Directory -Path $dir -Force)
 if($IncludeTrussJson){[void](New-Item -ItemType Directory -Path (Join-Path $dir 'trusses') -Force)}
 $byRule=[ordered]@{};$bySev=[ordered]@{}
 foreach($f in $o.findings){
  if($byRule.Contains($f.rule)){$byRule[$f.rule]++}else{$byRule[$f.rule]=1}
  if($bySev.Contains($f.sev)){$bySev[$f.sev]++}else{$bySev[$f.sev]=1}
 }
 $o.byRule=$byRule
 $cov=[ordered]@{compared=0;different=0;notCompared=0}
 if($o.dbOnly){
  $cov.notCompared=$o.counts.trusses*($Map.Count+2)
 }else{
  foreach($t in $o.trusses){$cov.compared+=$t.coverage.compared;$cov.different+=$t.coverage.different;$cov.notCompared+=$t.coverage.notCompared}
 }
 $o.coverage=$cov
 $summary=[ordered]@{project=$number;mode=$(if($Client){'Client'}else{'Server'});database=$Database;state=$st;archive=$o.archive;archiveTimestamp=$o.archiveTs;dbFilesCheckInId=$o.dbFilesCheckInId;projectValues=@([ordered]@{field='FilesCheckInId';db=$o.dbFilesCheckInId;file=$o.archiveTs;match=$(if($Client){$null}elseif($o.archiveTs -and $null -ne $o.dbFilesCheckInId){$o.dbFilesCheckInId -eq $o.archiveTs}else{$null});status=$(if($Client){'informational-only'}else{'compared'})});trusses=$o.counts.trusses;verdict=(VerdictOf $o $st);coverage=$cov;trussesWithRealMismatch=$o.counts.bad;trussesStaleDb=$o.counts.stale;pre_r7_db_only_trusses=$o.counts.pre;findingsByRule=$byRule;findingsBySeverity=$bySev}
 WriteText (Join-Path $dir 'summary.json') (ToJson $summary 4)
 $flat=[Collections.Generic.List[object]]::new()
 foreach($f in $o.findings){$flat.Add([ordered]@{ruleId=$f.rule;ruleName=(RuleName $f);rule=$Rules[$f.rule][1];project=$number;truss=$f.truss;field=$f.field;dbValue=$f.db;fileValue=$f.file;severity=$f.sev})}
 WriteText (Join-Path $dir 'findings.json') (ToJson $flat.ToArray() 4)
 $cc=[ordered]@{}
 foreach($f in $o.findings){$nm=RuleName $f;if($cc.Contains($nm)){$cc[$nm]++}else{$cc[$nm]=1}}
 $o.cats=$cc
 $chips=(($cc.Keys | Sort-Object | ForEach-Object {
  '<span class="chip" data-r="'+(Esc $_)+'">'+(Esc $_)+' ('+$cc[$_]+')</span>'
 }) -join '')
 if(-not $chips){$chips='no findings'}
 $headers=(($Groups | Where-Object {$_ -ne 'Version'} | ForEach-Object {'<th class=cell>'+$_+'</th>'}) -join '')
 $meta=$(if($o.archive){'archive '+(Esc $o.archive)}else{'no archive folder'})+' | DB FilesCheckInId '+(Esc $o.dbFilesCheckInId)+' | '+$o.counts.trusses+' trusses | '+$o.counts.bad+' real mismatch | '+$o.counts.stale+' stale-DB | '+$o.counts.pre+' before 2024.r7 DB-only (not counted)'+' | coverage '+$cov.compared+' compared, '+$cov.different+' different, '+$cov.notCompared+' not compared | <a href="../../contract.json">contract.json</a>'
 if($o.dbOnly -and -not $script:DbOnlyMetaJson){
  $dbMeta=[Collections.Generic.List[object]]::new()
  foreach($fld in $Map.Keys){$m=$MapMeta[$fld];$dbMeta.Add([object[]]@($fld,$m.group,$Map[$fld],$m.basis,$m.policy))}
  $script:DbOnlyMetaJson=ToJson $dbMeta.ToArray() 4 -compress
 }
 $dbMetaJson=if($o.dbOnly){$script:DbOnlyMetaJson}else{'[]'}
 $page=$ProjectTemplate.Replace('@@CSS@@',$Css).Replace('@@TITLE@@',(Esc $number)).Replace('@@BADGE@@',$Badge[$st]).Replace('@@STATE@@',$st).Replace('@@MODE@@',(Esc $modeLabel)).Replace('@@META@@',$meta).Replace('@@LEGEND@@',(CauseLegend)).Replace('@@CHIPS@@',$chips).Replace('@@HEADERS@@',$headers).Replace('@@GROUPS@@',(ToJson $Groups 2 -compress)).Replace('@@DBONLY@@',$(if($o.dbOnly){'true'}else{'false'})).Replace('@@DBMETA@@',$dbMetaJson)
 $split=$page.IndexOf('@@DATA@@',[StringComparison]::Ordinal)
 if($split -lt 0){throw 'Project template has no data marker.'}
 $tn=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
 $dbt=$null
 if($o.dbOnly){
  $dbt=DbTrusses $number
  $trussItems=@($dbt.Keys | Sort-Object)
 }else{
  $trussItems=@($o.trusses | Sort-Object @{Expression={if($_.findings.Count -gt 0){0}else{1}}},@{Expression={$_.name}})
 }
 $writer=[IO.StreamWriter]::new((Join-Path $dir 'project.html'),$false,[Text.UTF8Encoding]::new($false))
 try{
  $writer.Write($page.Substring(0,$split))
  $writer.Write('[')
  $written=0
  $chunk=[Collections.Generic.List[object]]::new()
  foreach($item in $trussItems){
   if($o.dbOnly){
     $dbRow=$dbt[$item][0]
     if($IncludeTrussJson){
      $t=[ordered]@{name=$item;cells=(Cells 'na');findings=[Collections.Generic.List[object]]::new();fields=[Collections.Generic.List[object]]::new();version=$null;pre=$false}
      FillDbOnlyFields $t $dbRow
     }else{$t=$null}
   }else{$t=$item}
    if($null -ne $t){
     $tf=[Collections.Generic.List[object]]::new()
     $fl=[Collections.Generic.List[object]]::new()
     foreach($f in $t.findings){
      $tf.Add([ordered]@{ruleId=$f.rule;ruleName=(RuleName $f);field=$f.field;db=$f.db;file=$f.file;severity=$f.sev;message=$f.msg})
      $fl.Add([object[]]@((RuleName $f),$f.field,$f.db,$f.file,$f.sev,$f.msg,$f.group))
     }
     if($IncludeTrussJson){
      $tj=[ordered]@{truss=$t.name;project=$number;version=$t.version;cells=$t.cells;pre_r7_db_only=[bool]$t.pre;findings=$tf.ToArray();fields=$t.fields.ToArray();coverage=$t.coverage}
      $tfile=SafeName $t.name;$k=2
      while(-not $tn.Add($tfile)){$tfile=(SafeName $t.name)+'_'+$k;$k++}
      WriteText (Join-Path $dir ('trusses\'+$tfile+'.json')) (ToJson $tj 50)
     }
   }
    if($o.dbOnly){
     $chunk.Add((CompactDbOnlyRow $item $dbRow))
    }else{
     $cellArr=[string[]]@($Groups | ForEach-Object {$t.cells[$_]})
     $ver=if($t.version){$t.version}else{@{}}
     $chunk.Add([object[]]@($t.name,$cellArr,$fl,$ver,$t.fields,$t.coverage))
    }
   if($chunk.Count -ge 128){$written=FlushProjectRows $writer $chunk $written}
  }
  $written=FlushProjectRows $writer $chunk $written
  $writer.Write(']')
  $writer.Write($page.Substring($split+'@@DATA@@'.Length))
 }finally{$writer.Dispose()}
 $indexRow=[ordered]@{project=$number;folder=$folder;state=$st;counts=$o.counts;coverage=$cov;cats=$cc;staleProject=[bool]$o.staleProject}
 [void]$results.Add($indexRow)
 if($TimingLog){
  $reportTimer.Stop()
  $serializeMs=1000.0*$script:ReportSerializeTicks/[Diagnostics.Stopwatch]::Frequency
  $writeMs=1000.0*$script:ReportWriteTicks/[Diagnostics.Stopwatch]::Frequency
  $htmlBytes=([IO.FileInfo]::new((Join-Path $dir 'project.html'))).Length
  LogTiming 'report' $reportTimer.Elapsed.TotalMilliseconds 0 $o.counts.trusses $serializeMs $writeMs $htmlBytes
 }
}
function FlushProjectRows($writer,$rows,[int]$written){
 if($rows.Count -eq 0){return $written}
 if($TimingLog){$chunkTimer=[Diagnostics.Stopwatch]::StartNew()}
 $json=(ToJson $rows.ToArray() 50 -compress).Replace('</','<\/')
 if($TimingLog){$chunkTimer.Stop();$script:ReportSerializeTicks+=$chunkTimer.ElapsedTicks;$chunkTimer.Restart()}
 if($written -gt 0){$writer.Write(',')}
 $writer.Write($json.Substring(1,$json.Length-2))
 if($TimingLog){$chunkTimer.Stop();$script:ReportWriteTicks+=$chunkTimer.ElapsedTicks}
 $written+=$rows.Count
 $rows.Clear()
 return $written
}
$results=[Collections.Generic.List[object]]::new()
$used=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$i=0;$batchEnd=-1
$progressTimer=[Diagnostics.Stopwatch]::StartNew()
try{
foreach($w in $work){
 if($TimingLog){$script:DiagnosticKind='archive';$script:DiagnosticProjectIndex=$i+1}
 if($i -gt $batchEnd){
  if($TimingLog){LogTiming 'batch_start';$batchTimer=[Diagnostics.Stopwatch]::StartNew()}
  $batchEnd=DbBatchEnd $work $i
  LoadDbBatch @($work[$i..$batchEnd] | ForEach-Object {$_[0]})
  if($TimingLog){$batchTimer.Stop();LogTiming 'batch' $batchTimer.Elapsed.TotalMilliseconds ($batchEnd-$i+1)}
 }
 $i++
 if($i -eq 1 -or $i -eq $work.Count -or $progressTimer.ElapsedMilliseconds -ge 10000){
  if($TimingLog){$progressCallTimer=[Diagnostics.Stopwatch]::StartNew()}
  Write-Progress -Activity 'Analyzing projects' -Status ($w[0]+' ('+$i+' of '+$work.Count+')') -PercentComplete (100*$i/$work.Count)
  if($TimingLog){$progressCallTimer.Stop();LogTiming 'progress' $progressCallTimer.Elapsed.TotalMilliseconds}
  $progressTimer.Restart()
 }
 $number=$w[0]
 try{
  if($TimingLog){LogTiming 'analysis_start';$analysisTimer=[Diagnostics.Stopwatch]::StartNew()}
  try{$o=AnalyzeProject $number $w[1] $cut}
  catch{Write-Warning ($number+": "+$_.Exception.Message+" @ "+$_.InvocationInfo.ScriptLineNumber);continue}
  if($TimingLog){$analysisTimer.Stop();LogTiming 'analysis' $analysisTimer.Elapsed.TotalMilliseconds 0 $o.counts.trusses}
  WriteProject $o $number
 }finally{ReleaseProject $number;$o=$null;LogTiming 'project_released'}
}
if($IncludeDbOnly){
 if($TimingLog){$script:DiagnosticKind='db_only';$script:DiagnosticProjectIndex=0;LogTiming 'db_only_discovery_start'}
 $discoveryWatch=[Diagnostics.Stopwatch]::StartNew()
 $folderScanTimer=[Diagnostics.Stopwatch]::StartNew()
 $have=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
 foreach($d in (Get-ChildItem -LiteralPath $root -Directory)){[void]$have.Add($d.Name)}
 $folderScanTimer.Stop()
 LogTiming 'folder_scan' $folderScanTimer.Elapsed.TotalMilliseconds $have.Count
 $dbOnly=[Collections.Generic.List[object]]::new()
 $dbCommand=$script:Conn.CreateCommand()
 $dbCommand.CommandText='SELECT p.ProjectNumber,p.FilesCheckInId FROM dbo.Project p ORDER BY p.ProjectNumber'
 $dbCommand.CommandTimeout=120
 $dbReader=$null;$stoppedEarly=$false
 if($TimingLog){LogTiming 'sql_discovery_start';$sqlDiscoveryTimer=[Diagnostics.Stopwatch]::StartNew()}
 try{
  $dbReader=$dbCommand.ExecuteReader()
  while($dbReader.Read()){
   $pn=[string]$dbReader.GetValue(0)
   if($have.Contains($pn)){continue}
   if($ProjectId -and ($ProjectId -notcontains $pn)){continue}
   $ci=if($dbReader.IsDBNull(1)){$null}else{[string]$dbReader.GetValue(1)}
   $dbOnly.Add(@($pn,$ci))
   if($Limit -gt 0 -and $dbOnly.Count -ge $Limit){$stoppedEarly=$true;break}
  }
 }finally{
  if($stoppedEarly){$dbCommand.Cancel()}
  if($dbReader){$dbReader.Dispose()}
  $dbCommand.Dispose()
 }
 if($TimingLog){$sqlDiscoveryTimer.Stop();LogTiming 'sql_discovery' $sqlDiscoveryTimer.Elapsed.TotalMilliseconds $dbOnly.Count}
 $discoveryWatch.Stop()
 LogTiming 'db_only_discovery' $discoveryWatch.Elapsed.TotalMilliseconds $dbOnly.Count
 Write-Verbose ('DB-only discovery: {0} projects in {1:N2}s' -f $dbOnly.Count,$discoveryWatch.Elapsed.TotalSeconds)
 Write-Host ('DB-only projects: '+$dbOnly.Count)
 $k=0;$batchEnd=-1
 $dbProgressTimer=[Diagnostics.Stopwatch]::StartNew()
 foreach($w in $dbOnly){
  if($TimingLog){$script:DiagnosticProjectIndex=$k+1}
  if($k -gt $batchEnd){
   if($TimingLog){LogTiming 'batch_start'}
   $batchWatch=[Diagnostics.Stopwatch]::StartNew()
   $batchEnd=DbBatchEnd $dbOnly $k
   LoadDbBatch @($dbOnly[$k..$batchEnd] | ForEach-Object {$_[0]})
   $batchWatch.Stop()
   LogTiming 'batch' $batchWatch.Elapsed.TotalMilliseconds ($batchEnd-$k+1)
   Write-Verbose ('DB-only batch: {0} projects loaded in {1:N2}s' -f ($batchEnd-$k+1),$batchWatch.Elapsed.TotalSeconds)
  }
  $k++
  if($k -eq 1 -or $k -eq $dbOnly.Count -or $dbProgressTimer.ElapsedMilliseconds -ge 10000){
   if($TimingLog){$progressCallTimer=[Diagnostics.Stopwatch]::StartNew()}
   Write-Progress -Activity 'DB-only projects' -Status ($w[0]+' ('+$k+' of '+$dbOnly.Count+')') -PercentComplete (100*$k/$dbOnly.Count)
   if($TimingLog){$progressCallTimer.Stop();LogTiming 'progress' $progressCallTimer.Elapsed.TotalMilliseconds}
   $dbProgressTimer.Restart()
  }
  $projectWatch=[Diagnostics.Stopwatch]::StartNew()
  try{
   if($TimingLog){LogTiming 'analysis_start';$analysisTimer=[Diagnostics.Stopwatch]::StartNew()}
   try{$o=AnalyzeDbOnly $w[0] $w[1]}
   catch{Write-Warning ($w[0]+": "+$_.Exception.Message);continue}
   if($TimingLog){$analysisTimer.Stop();LogTiming 'analysis' $analysisTimer.Elapsed.TotalMilliseconds 0 $o.counts.trusses}
   $o.state='DB only'
   WriteProject $o $w[0]
   $projectWatch.Stop()
   if($VerbosePreference -eq 'Continue'){
    $heapMB=[Math]::Round([GC]::GetTotalMemory($false)/1MB,1)
    Write-Verbose ('DB-only {0}: {1} trusses, report {2:N2}s, managed heap {3:N1} MB' -f $w[0],$o.counts.trusses,$projectWatch.Elapsed.TotalSeconds,$heapMB)
   }
  }finally{ReleaseProject $w[0];$o=$null;LogTiming 'project_released'}
 }
 Write-Progress -Activity 'DB-only projects' -Completed
}
}finally{LoadDbBatch @();$script:Conn.Dispose();Write-Progress -Activity 'Analyzing projects' -Completed}

if($TimingLog){$script:DiagnosticKind='index';$script:DiagnosticProjectIndex=0;LogTiming 'index_start';$indexTimer=[Diagnostics.Stopwatch]::StartNew()}
$order=@{'Inconsistent'=0;'Stale DB'=1;'Not in DB'=2;'Consistent'=3;'Before 2024.r7'=4;'DB only'=5}
$ranked=@($results | Sort-Object @{Expression={$order[$_.state]}},@{Expression={-($_.counts.bad+$_.counts.stale)}},@{Expression={$_.project}})
$cnt=@{};foreach($s in $order.Keys){$cnt[$s]=0};foreach($o in $ranked){$cnt[$o.state]++}
$tabNames=@('Inconsistent','Stale DB','Not in DB','Consistent','Before 2024.r7')
if($IncludeDbOnly){$tabNames+='DB only'}
$tabCount=$cnt
$tabHints=@{'Inconsistent'='Files and database disagree';'Stale DB'='Database older than archive';'Not in DB'='Archive has no project row';'Consistent'='No differences in compared values';'Before 2024.r7'='Database is the source';'DB only'='No archive folder'}
$tabs=(($tabNames | ForEach-Object {'<div class="tab'+$(if($_ -eq 'Inconsistent'){' on'}else{''})+'" data-t="'+$_+'"><b>'+$tabCount[$_]+'</b>'+$_+'<br><span>'+$tabHints[$_]+'</span></div>'}) -join '')
$allChips=(((CauseDefinitions).GetEnumerator() | ForEach-Object {'<span class="chip" data-c="'+(Esc $_.Key)+'" title="'+(Esc $_.Value)+'">'+(Esc $_.Key)+'</span>'}) -join '')
$sb=[Text.StringBuilder]::new()
foreach($o in $ranked){
 $top=(($o.cats.GetEnumerator() | Sort-Object @{Expression={-$_.Value}},@{Expression={$_.Key}} | Select-Object -First 2 | ForEach-Object {'<span class="badge b-pre">'+(Esc $_.Key)+' '+$_.Value+'</span>'}) -join '')
 [void]$sb.Append('<tr data-s="'+$o.state+'" data-c="'+(Esc ($o.cats.Keys -join '|'))+'"><td>'+$(if($o.folder){'<a href="projects/'+(Esc $o.folder)+'/project.html">'+(Esc $o.project)+'</a>'}else{Esc $o.project})+'</td><td><span class="badge '+$Badge[$o.state]+'">'+$o.state+'</span></td><td>'+(Esc (VerdictOf $o $o.state))+'</td><td>'+$top+'</td><td>'+$o.counts.trusses+'</td></tr>')
}
$agg=[ordered]@{compared=0;different=0;notCompared=0}
foreach($o in $ranked){$agg.compared+=$o.coverage.compared;$agg.different+=$o.coverage.different;$agg.notCompared+=$o.coverage.notCompared}
$contractFields=[Collections.Generic.List[object]]::new()
foreach($k in $Map.Keys){
 $m=$MapMeta[$k]
 $contractFields.Add([ordered]@{field=$k;column=$Map[$k];table=$(if($k -in @('Weight','IsSymmetric')){'ComponentHeader'}elseif($k -eq 'Quantity'){'Component'}else{'ComponentTruss'});group=$m.group;policy=$m.policy;basis=$m.basis;status=$(if($m.basis -eq 'candidate'){'candidate (not confirmed)'}else{'mapped'})})
}
$contract=[ordered]@{
 name='csdirector-database-project-insights contract';
 version=1;
 generated=(Get-Date).ToString('yyyy-MM-dd HH:mm:ss');
 script=$PSCommandPath;
 mode=$(if($Client){'Client'}else{'Server'});
 database=$Database;
 projectsRoot=$root;
 archive='latest .csdproj per project folder (or expanded project folder on Client mode)';
 sourceTables=[ordered]@{
  'Project'='p.ProjectNumber, p.FilesCheckInId; check-in compared against archive timestamp on Server mode, informational-only on Client mode';
  'ComponentHeader'='h.Name, h.TrussFileCheckSum, h.WeightPounds, h.IsSymmetrical; name identity, .tdlTruss MD5, Weight and IsSymmetric';
  'Component'='c.ComponentQuantity; mapped as Quantity';
  'ComponentTruss'='Explicit columns only: AnalysisVersion, CompatibilityVersion and mapped columns; unmapped payloads are not fetched';
  'other'='full contract also involves DB tables not yet queried'
 };
 groups=$Groups;
 policies=[ordered]@{
  numeric='abs difference <= 0.001 applied only when both sides parse as numbers; otherwise ordinal case-sensitive text';
  integer='exact integer equality for quantity/ply/count values; ordinal text fallback';
  boolean='bool or numeric 0/1 only; never NULL-to-false coercion; ordinal text fallback';
  text='ordinal case-sensitive exact string equality';
  nulls='missing values are never matches; missing field, absent DB column, database NULL, file NULL and missing side are distinguished; a both-null pair is unknown, not true and distinct from both sides absent'
 };
 identityContract=[ordered]@{
  name='Truss name: Trusses/<name>.tdlTruss controls membership against ComponentHeader.Name; Trusses/generated/<name>.json is comparison data only and is read only when the .tdlTruss exists';
  checksum='.tdlTruss MD5 <-> ComponentHeader.TrussFileCheckSum';
  version='JSON Version <-> ComponentTruss.AnalysisVersion';
  checkIn='Project.FilesCheckInId <-> archive timestamp; informational-only on Client mode (archive timestamp is not compared there)'
 };
 mappedFields=$contractFields.ToArray();
 candidates=[ordered]@{note='mapped rows with basis=candidate are unresolved: names/meaning not confirmed, values retained, never flagged as mismatches';fields=@($Map.Keys|Where-Object{$MapMeta[$_].basis -eq 'candidate'})};
 nestedPending=[ordered]@{
  'CarriedTrusses'='structured encoding pending; not string-compared';
  'JigData'='structured encoding pending; not string-compared';
  'other nested objects/arrays'='not included in output; relational mapping pending; not raw-text-compared'
 };
 lookupPending='Unmapped ComponentTruss FK columns (*Key) are not fetched or compared to human names; lookup table identities pending';
 coverage='compared/different/notCompared for mapped fields, version and checksum only; unmapped inventory excluded; values not compared are not claims of equality'
}
WriteText (Join-Path $outDir 'contract.json') (ToJson $contract 6)
$index=$IndexTemplate.Replace('@@CSS@@',$Css).Replace('@@MODE@@',(Esc $modeLabel)).Replace('@@COUNT@@',[string]$ranked.Count).Replace('@@WHEN@@',(Get-Date).ToString('yyyy-MM-dd HH:mm')).Replace('@@SOURCE@@',(Esc $root)).Replace('@@AGGC@@',('compared '+$agg.compared+', different '+$agg.different+', not compared '+$agg.notCompared)).Replace('@@TABS@@',$tabs).Replace('@@CHIPS@@',$allChips).Replace('@@ROWS@@',$sb.ToString()).Replace('@@LEGEND@@',(CauseLegend))
WriteText (Join-Path $outDir 'index.html') $index
if($TimingLog){$indexTimer.Stop();LogTiming 'index' $indexTimer.Elapsed.TotalMilliseconds $ranked.Count;LogTiming 'complete'}
Write-Host ''
Write-Host ('Projects: '+$ranked.Count)
foreach($s in $tabNames){Write-Host ('  '+$s+': '+$cnt[$s])}
$reportUri=([System.Uri]([IO.Path]::GetFullPath((Join-Path $outDir 'index.html')))).AbsoluteUri
$escape=[char]27
Write-Host ("Report: ${escape}]8;;${reportUri}${escape}\${reportUri}${escape}]8;;${escape}\")

#requires -Version 5.1
<#
.SYNOPSIS
Read-only comparison of CSDirector database trusses with the TrussStudio files inside .csdproj archives. Produces a summary index, one page per project and JSON for each project and truss.
.EXAMPLE
.\csdirector-database-project-insights.ps1 -SqlUser insights_ro -Server 192.168.71.183 -ProjectsRoot C:\temp\csdproj-samples
.EXAMPLE
.\csdirector-database-project-insights.ps1 -ProjectId S-789,I-4028
.EXAMPLE
.\csdirector-database-project-insights.ps1 -Limit 5
.NOTES
With -SqlUser the password is read only from the CSD_SQL_PASSWORD environment variable. Archives and SQL records are never changed.
Projects whose trusses are all older than -LegacyCutoffVersion (pre 2024.r7) are DB-only and are not counted as inconsistent.
A project whose DB check-in id is older than the archive timestamp is reported as Stale DB rather than Inconsistent.
#>
[CmdletBinding()]
param(
 [string]$ProjectsRoot='C:\SST\Server\Projects',
 [string]$Server='.\SST',
 [string]$SqlUser,
 [string]$Database='CSDatabaseServer',
 [string]$OutputDirectory=(Join-Path (Get-Location).Path 'output-insights'),
 [string[]]$ProjectId,
 [ValidatePattern('^\d+(\.\d+){0,3}$')][string]$LegacyCutoffVersion='2024.7.2.7',
 [ValidateRange(0,2147483647)][int]$Limit=0
)
$ErrorActionPreference='Stop'
Write-Host ''
Write-Host 'CSDirector: Database and Project file Insights' -ForegroundColor Cyan
Write-Host ('Compares DB trusses with TrussStudio files in each .csdproj using '+$(if($SqlUser){'SQL authentication ('+$SqlUser+').'}else{'Windows authentication.'}))
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
$Groups=@('Identity','Version','Geometry','Pitch','Flags')
$Map=[ordered]@{Quantity='ComponentQuantity';NumPlies='TrussPlyCount';Thickness='TrussThicknessInches';OverallTrussHeight='OverallTrussHeightInches';LeftOverhang='LeftOverhangInches';RightOverhang='RightOverhangInches';LeftCantilever='LeftCantileverInches';RightCantilever='RightCantileverInches';LeftHeelHeight='LeftHeelHeightInches';RightHeelHeight='RightHeelHeightInches';PitchLeftTopOver12='LeftTopPitchOver12';PitchRightTopOver12='RightTopPitchOver12';PitchLeftBottomOver12='LeftBottomPitchOver12';PitchRightBottomOver12='RightBottomPitchOver12';IsAttic='IsAttic';IsGable='IsGable';IsGirder='IsGirder';IsFlipped='IsFlipped';OCSpacing='OCSpacing'}
$FieldGroup=@{Quantity='Identity';PitchLeftTopOver12='Pitch';PitchRightTopOver12='Pitch';PitchLeftBottomOver12='Pitch';PitchRightBottomOver12='Pitch';IsAttic='Flags';IsGable='Flags';IsGirder='Flags';IsFlipped='Flags'}
$Rules=[ordered]@{
 'R-PROJ'=@('Project not in DB','Archive exists but no Project row with this number in the DB.')
 'R-FCI'=@('DB differs from archive','DB check-in (Project.FilesCheckInId) should equal the timestamp in the archive file name.')
 'R-FILE'=@('Truss missing from archive','DB truss at a current (r7+) version has no Trusses/generated/<name>.json in the archive. Pre-r7 DB-only trusses are expected and excluded.')
 'R-DB'=@('Truss missing from DB','Truss file exists in the archive but no ComponentHeader row with that name.')
 'R-DUP'=@('Duplicate truss in DB','More than one DB row for the same truss name.')
 'R-CKSUM'=@('Truss file checksum differs','MD5 of Trusses/<name>.tdlTruss should equal ComponentHeader.TrussFileCheckSum.')
 'R-TDL'=@('Truss file missing','No .tdlTruss to checksum.')
 'R-VER'=@('Truss version differs','Truss JSON Version should equal ComponentTruss.AnalysisVersion. Reported as stale when the DB is frozen pre-r7 and the file is current.')
 'R-FLD'=@('Truss value differs','Mapped truss JSON field should equal its ComponentTruss/Component column (tolerance 0.001).')
}
$StaleRules=@('R-CKSUM','R-DB','R-FLD','R-VER','R-FILE')
$Badge=@{'Inconsistent'='b-bad';'Consistent'='b-ok';'Pre-r7'='b-pre';'Stale DB'='b-warn';'Not in DB'='b-warn';'DB only'='b-pre'}

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
function Cells($v){$c=[ordered]@{};foreach($g in $Groups){$c[$g]=$v};$c}
function IdentityOnly($state){$c=Cells 'na';$c['Identity']=$state;$c}
function Q($sql,$params){
 $cmd=$script:Conn.CreateCommand();$cmd.CommandText=$sql;$cmd.CommandTimeout=120
 foreach($k in $params.Keys){[void]$cmd.Parameters.AddWithValue('@'+$k,$params[$k])}
 $dt=[Data.DataTable]::new()
 [void][Data.SqlClient.SqlDataAdapter]::new($cmd).Fill($dt)
 return ,$dt
}
function AddF($row,$All,$stale,$rule,$field,$dbv,$filev,$msg,$grp,$sev='error'){
 if($StaleRules -contains $rule -and ($stale -or $sev -eq 'warn')){$sev='stale'}
 $item=[ordered]@{rule=$rule;truss=$row.name;field=$field;db=$dbv;file=$filev;sev=$sev;msg=$msg;group=$grp}
 $row.findings.Add($item);$All.Add($item)
 $nv=if($sev -eq 'stale'){'stale'}else{'bad'}
 $c=$row.cells[$grp]
 if($c -ne 'bad' -and ($nv -eq 'bad' -or $c -eq 'ok')){$row.cells[$grp]=$nv}
}
function ReadEntryBytes($entry){
 $s=$entry.Open();$ms=[IO.MemoryStream]::new()
 try{$s.CopyTo($ms)}finally{$s.Dispose()}
 return ,$ms.ToArray()
}
function ReadEntryJson($entry){
 $s=$entry.Open();$r=[IO.StreamReader]::new($s,[Text.Encoding]::UTF8,$true)
 try{$t=$r.ReadToEnd()}finally{$r.Dispose()}
 $t | ConvertFrom-Json
}
function DbTrusses($number){
 $dt=Q 'SELECT h.Name,h.TrussFileCheckSum,c.ComponentQuantity,t.* FROM dbo.Project p JOIN dbo.ComponentHeader h ON h.ProjectEventKey=p.EventKey LEFT JOIN dbo.Component c ON c.ComponentKey=h.ComponentHeaderKey LEFT JOIN dbo.ComponentTruss t ON t.ComponentKey=c.ComponentKey WHERE p.ProjectNumber=@p ORDER BY h.Name' @{p=$number}
 $d=[Collections.Generic.Dictionary[string,object]]::new()
 foreach($r in $dt.Rows){
  $o=@{}
  foreach($col in $dt.Columns){
   $v=$r[$col.ColumnName]
   if($v -is [DBNull]){$o[$col.ColumnName]=$null}
   elseif($v -is [byte[]]){if($col.ColumnName -eq 'TrussFileCheckSum'){$o[$col.ColumnName]=[BitConverter]::ToString($v)}}
   elseif($v -is [string] -and $v.Length -gt 200){continue}
   elseif($v -is [guid] -or $v -is [datetime]){$o[$col.ColumnName]=[string]$v}
   else{$o[$col.ColumnName]=$v}
  }
  $n=[string]$o['Name']
  if(-not $d.ContainsKey($n)){$d[$n]=[Collections.Generic.List[object]]::new()}
  $d[$n].Add($o)
 }
 return ,$d
}
function AnalyzeProject($number,$archivePath,$cut){
 $m=[regex]::Match([IO.Path]::GetFileName($archivePath),'_(\d{14})')
 $ts=if($m.Success){$m.Groups[1].Value}else{$null}
 $pdt=Q 'SELECT ProjectNumber,FilesCheckInId FROM dbo.Project WHERE ProjectNumber=@p' @{p=$number}
 $hasProject=$pdt.Rows.Count -gt 0
 $dbCheckIn=$null
 if($hasProject -and $pdt.Rows[0]['FilesCheckInId'] -isnot [DBNull]){$dbCheckIn=[string]$pdt.Rows[0]['FilesCheckInId']}
 $dbt=[Collections.Generic.Dictionary[string,object]]::new()
 if($hasProject){$dbt=DbTrusses $number}
 $out=[ordered]@{project=$number;archive=[IO.Path]::GetFileName($archivePath);archiveTs=$ts;dbFilesCheckInId=$dbCheckIn;notInDb=(-not $hasProject);trusses=[Collections.Generic.List[object]]::new();findings=[Collections.Generic.List[object]]::new()}
 $F=$out.findings
 if(-not $hasProject){
  $F.Add([ordered]@{rule='R-PROJ';truss=$null;field='Project';db=$null;file=$number;sev='error';msg='Project not in DB';group='Project'})
 }elseif($ts -and $dbCheckIn -ne $ts){
  $F.Add([ordered]@{rule='R-FCI';truss=$null;field='FilesCheckInId';db=$dbCheckIn;file=$ts;sev='warn';msg='DB check-in id differs from archive timestamp';group='Project'})
 }
 $stale=[bool]($hasProject -and $ts -and $null -ne $dbCheckIn -and [string]::CompareOrdinal($dbCheckIn,$ts) -lt 0)
 $zip=[IO.Compression.ZipFile]::OpenRead($archivePath)
 try{
  $files=[Collections.Generic.Dictionary[string,object]]::new()
  $tdl=[Collections.Generic.Dictionary[string,object]]::new()
  foreach($e in $zip.Entries){
   $mm=[regex]::Match($e.FullName,'^Trusses/generated/(.+)\.json$')
   if($mm.Success){$files[$mm.Groups[1].Value]=$e;continue}
   $mm=[regex]::Match($e.FullName,'^Trusses/(.+)\.tdlTruss$')
   if($mm.Success){$tdl[$mm.Groups[1].Value]=$e}
  }
  $names=[Collections.Generic.List[string]]::new()
  foreach($k in $files.Keys){$names.Add($k)}
  foreach($k in $dbt.Keys){if(-not $files.ContainsKey($k)){$names.Add($k)}}
  $names.Sort([StringComparer]::Ordinal)
  foreach($name in $names){
   $row=[ordered]@{name=$name;cells=(Cells 'ok');findings=[Collections.Generic.List[object]]::new();fields=[Collections.Generic.List[object]]::new();version=$null;pre=$false}
   $rows=$null
   if($dbt.ContainsKey($name)){$rows=$dbt[$name]}
   if(-not $files.ContainsKey($name)){
    $av=$null
    if($rows){$av=ParseVer $rows[0]['AnalysisVersion']}
    if($rows -and $null -ne $av -and (VerLess $av $cut)){
     $row.cells=Cells 'pre';$row.pre=$true
     $row.version=[ordered]@{file=$null;db=$rows[0]['AnalysisVersion']}
     $out.trusses.Add($row);continue
    }
    AddF $row $F $stale 'R-FILE' ('Trusses/generated/'+$name+'.json') 'present' 'missing' 'Truss in DB but no file in archive' 'Identity'
    $row.cells=IdentityOnly $row.cells['Identity']
    $out.trusses.Add($row);continue
   }
   $j=ReadEntryJson $files[$name]
   if(-not $rows){
    AddF $row $F $stale 'R-DB' 'ComponentHeader.Name' 'missing' $name 'Truss in file but not in DB' 'Identity'
    $row.cells=IdentityOnly $row.cells['Identity']
    $out.trusses.Add($row);continue
   }
   if($rows.Count -gt 1){AddF $row $F $stale 'R-DUP' 'ComponentHeader.Name' $rows.Count 1 'Duplicate truss name rows in DB' 'Identity'}
   $r=$rows[0]
   if($tdl.ContainsKey($name)){
    $md5=[Security.Cryptography.MD5]::Create()
    try{$h=[BitConverter]::ToString($md5.ComputeHash((ReadEntryBytes $tdl[$name])))}finally{$md5.Dispose()}
    if($null -ne $r['TrussFileCheckSum'] -and ([string]$r['TrussFileCheckSum']).ToUpperInvariant() -ne $h){
     AddF $row $F $stale 'R-CKSUM' 'TrussFileCheckSum' ([string]$r['TrussFileCheckSum']) $h '.tdlTruss MD5 differs from DB checksum' 'Identity'
    }
   }else{
    AddF $row $F $stale 'R-TDL' ('Trusses/'+$name+'.tdlTruss') 'n/a' 'missing' 'No .tdlTruss for truss' 'Identity' 'warn'
   }
   $jv=$j.PSObject.Properties['Version'];$jv=if($jv){$jv.Value}else{$null}
   $jc=$j.PSObject.Properties['CompatibilityVersion'];$jc=if($jc){$jc.Value}else{$null}
   $dv=$r['AnalysisVersion']
   $fvp=ParseVer $jv;$dvp=ParseVer $dv
   $row.version=[ordered]@{file=$jv;db=$dv;compat_file=$jc;compat_db=$r['CompatibilityVersion']}
   $same=([string]$jv -eq [string]$dv)
   if(-not $same){
    if($null -ne $dvp -and (VerLess $dvp $cut) -and -not ($null -ne $fvp -and (VerLess $fvp $cut))){
     AddF $row $F $stale 'R-VER' 'Version vs AnalysisVersion' $dv $jv 'DB frozen at pre-r7 version; file is current (re-opened in Studio, DB not refreshed)' 'Version' 'warn'
    }else{
     AddF $row $F $stale 'R-VER' 'Version vs AnalysisVersion' $dv $jv 'File and DB version differ' 'Version'
    }
   }
   $row.fields.Add([ordered]@{field='Version';group='Version';column='AnalysisVersion';db=$dv;file=$jv;match=$same})
   foreach($fld in $Map.Keys){
    $col=$Map[$fld]
    $grp=if($FieldGroup.ContainsKey($fld)){$FieldGroup[$fld]}else{'Geometry'}
    $pa=$j.PSObject.Properties[$fld];$a=if($pa){$pa.Value}else{$null}
    $b=$r[$col]
    if($null -eq $a -or $null -eq $b){
     $row.fields.Add([ordered]@{field=$fld;group=$grp;column=$col;db=$b;file=$a;match=$null});continue
    }
    if($a -is [bool] -or $b -is [bool]){$eq=([bool]$a -eq [bool]$b)}
    else{
     $x=0.0;$y=0.0
     if([double]::TryParse(([string]$a),[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$x) -and [double]::TryParse(([Convert]::ToString($b,[Globalization.CultureInfo]::InvariantCulture)),[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$y)){$eq=([math]::Abs($x-$y) -le 0.001)}
     else{$eq=([string]$a -eq [string]$b)}
    }
    $row.fields.Add([ordered]@{field=$fld;group=$grp;column=$col;db=$b;file=$a;match=$eq})
    if(-not $eq){AddF $row $F $stale 'R-FLD' $fld $b $a ($col+' differs') $grp}
   }
   $out.trusses.Add($row)
  }
 }finally{$zip.Dispose()}
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
 if($o.notInDb){return 'Not in DB'}
 foreach($f in $o.findings){if($f.sev -eq 'error'){return 'Inconsistent'}}
 if($o.counts.trusses -gt 0 -and $o.counts.pre -eq $o.counts.trusses){return 'Pre-r7'}
 if($o.findings.Count -gt 0){return 'Stale DB'}
 'Consistent'
}
function VerdictOf($o,$st){
 $n=$o.counts.trusses
 switch($st){
  'Consistent'{"All $n trusses match."}
  'Not in DB'{'Archive exists, but the project is not in the DB'}
  'DB only'{'In the DB, but no archive folder on disk. Not compared.'}
  'Pre-r7'{'Frozen pre-r7 DB state; no Studio files to compare.'}
  'Stale DB'{"DB older than the archive (re-saved in Studio). $($o.counts.stale) of $n trusses differ."}
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
.tabs{display:flex;gap:2px;border-bottom:2px solid #c9d2db;margin-bottom:10px}.tab{padding:7px 14px;cursor:pointer;background:#e4e9ee;border-radius:6px 6px 0 0}.tab.on{background:#fff;font-weight:600;border:1px solid #c9d2db;border-bottom:2px solid #fff;margin-bottom:-2px}
input[type=text]{padding:6px 9px;border:1px solid #b8c2cc;border-radius:4px;width:260px}
table{border-collapse:collapse;background:#fff;width:100%;box-shadow:0 1px 2px #0002}th,td{padding:6px 10px;border-bottom:1px solid #e3e8ed;text-align:left;vertical-align:top}th{background:#eef1f4;position:sticky;top:0}
.chip{display:inline-block;background:#eef1f4;border-radius:4px;padding:1px 6px;margin:1px;font-size:12px;cursor:pointer;border:1px solid transparent}.chip.on{border-color:#0b5cad;background:#dcebfa}
.cell{width:90px;text-align:center;white-space:nowrap;cursor:pointer;font-weight:700}.c-ok{background:#d6f1de;color:#0d5c28}.c-bad{background:#f4b9b3;color:#8a160c}.c-pre{background:#d5d9de;color:#4a5560}.c-stale{background:#fdebc8;color:#7a4b00}.c-na{background:#f3f3f3;color:#aaa}
.legend span{margin-right:14px;font-size:12px}.sw{display:inline-block;width:12px;height:12px;vertical-align:-1px;margin-right:4px;border-radius:2px}
.sbs td{font-family:Consolas,monospace;font-size:12px;word-break:break-all}.sbs .d{background:#fbe3e0}.sbs .f{background:#e3f1e6}
.dd td{font-weight:600}.note{background:#fff8e1;border-left:4px solid #e0a800;padding:6px 10px;margin:8px 0;font-size:13px}
'@

$ProjectTemplate=@'
<!doctype html><meta charset=utf-8><title>@@TITLE@@</title><style>@@CSS@@</style>
<header><h1><a href="../../index.html" style="color:#9cd0ff">All projects</a> / @@TITLE@@ <span class="badge @@BADGE@@">@@STATE@@</span></h1>
<small>@@META@@</small></header>
<main><div class="legend"><span><i class="sw c-ok"></i>matches</span><span><i class="sw c-bad"></i>real mismatch</span><span><i class="sw c-stale"></i>stale DB (file newer)</span><span><i class="sw c-pre"></i>pre-r7 DB-only (expected)</span><span><i class="sw c-na"></i>not comparable</span></div>
<p>Filter by issue: @@CHIPS@@</p>
<p><input type=text id=q placeholder="filter truss name"> <label><input type=checkbox id=onlyBad checked> only trusses with findings</label> <span id=cnt></span></p>
<table style="table-layout:fixed"><thead><tr><th>Truss</th>@@HEADERS@@<th class=cell style="width:220px">Version (file / DB)</th></tr></thead><tbody id=tb></tbody></table></main>
<script>
const G=@@GROUPS@@,D=@@DATA@@,CL={ok:'c-ok',bad:'c-bad',stale:'c-stale',pre:'c-pre',na:'c-na'},SY={ok:'&#10003;',bad:'&#8800;',stale:'&#8635;',pre:'&#9208;',na:'&#8709;'},open_=new Set(),all_=new Set();
let rules=new Set();
function esc(s){return String(s==null?'':s).replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]))}
function fv(x){return x==null?'(null)':esc(x)}
function det(i){const r=D[i],F=r[4]||[];let h='<tr><td colspan=6 style="background:#fafbfc">';
if(r[2].length)h+='<p class=note>'+r[2].map(f=>'<b>'+f[6]+'</b>: '+f[0]+' - '+esc(f[5])+(f[4]=='stale'?' (stale DB)':'')).filter((v,k,a)=>a.indexOf(v)==k).join('<br>')+'</p>';
const bad=F.filter(f=>f.match===false),rest=F.filter(f=>f.match!==false);
const row=(f)=>'<tr class="'+(f.match===false?'dd':'')+'"><td>'+f.group+'</td><td>'+esc(f.field)+'</td><td>'+esc(f.column)+'</td><td class="'+(f.match===false?'d':f.match?'f':'')+'">'+fv(f.db)+'</td><td class="'+(f.match===false?'d':f.match?'f':'')+'">'+fv(f.file)+'</td><td>'+(f.match===false?'differs':f.match?'match':'n/a')+'</td></tr>';
const extra=r[2].filter(f=>f[0]!='Truss value differs'&&f[0]!='Truss version differs');
const xr=extra.map(f=>'<tr class=dd><td>'+f[6]+'</td><td>'+esc(f[1])+'</td><td>-</td><td class=d>'+fv(f[2])+'</td><td class=d>'+fv(f[3])+'</td><td>'+(f[4]=='stale'?'stale':'differs')+'</td></tr>').join('');
if(F.length||xr){h+='<table class=sbs><tr><th>Group</th><th>JSON field</th><th>DB column</th><th>DB</th><th>File</th><th></th></tr>'+xr+bad.map(row).join('')+(all_.has(i)?rest.map(row).join(''):'')+'</table>';
if(rest.length)h+='<a href="#" onclick="all_.has('+i+')?all_.delete('+i+'):all_.add('+i+');draw();return false">'+(all_.has(i)?'hide':'show')+' '+rest.length+' matching/other fields</a>';}
else h+='<p>No comparable fields (truss present on one side only or pre-r7 DB-only).</p>';
return h+'</td></tr>'}
function draw(){const q=document.getElementById('q').value.toLowerCase(),ob=document.getElementById('onlyBad').checked;let n=0,h='';
D.forEach((r,i)=>{if(ob&&!r[2].length)return;if(q&&r[0].toLowerCase().indexOf(q)<0)return;if(rules.size&&!r[2].some(f=>rules.has(f[0])))return;n++;if(n>400)return;
const v=r[3];h+='<tr style="cursor:pointer" onclick="tg('+i+')"><td>'+(open_.has(i)?'&#9662; ':'&#9656; ')+esc(r[0])+'</td>'+r[1].map((c,k)=>G[k]=='Version'?'':'<td class="cell '+CL[c]+'">'+SY[c]+'</td>').join('')+'<td class="cell '+CL[r[1][1]]+'">'+SY[r[1][1]]+(v.file||v.db?' '+esc(v.file||'-')+' / '+esc(v.db||'-'):'')+'</td></tr>'+(open_.has(i)?det(i):'')});
document.getElementById('tb').innerHTML=h;document.getElementById('cnt').textContent=n+' trusses'+(n>400?' (first 400 shown)':'')}
function tg(i){open_.has(i)?open_.delete(i):open_.add(i);draw()}
document.getElementById('q').oninput=draw;document.getElementById('onlyBad').onchange=draw;
document.querySelectorAll('.chip').forEach(c=>c.onclick=()=>{const r=c.dataset.r;rules.has(r)?rules.delete(r):rules.add(r);c.classList.toggle('on');draw()});
draw();
</script>
'@

$IndexTemplate=@'
<!doctype html><meta charset=utf-8><title>CSDirector DB vs csdproj insights</title><style>@@CSS@@</style>
<header><h1>CSDirector DB vs .csdproj insights <span class="badge b-warn">@@COUNT@@ projects</span></h1><small>Summary only. Open a project for the truss-by-truss field comparison. Pre-r7 DB-only projects are not counted as inconsistent; Stale DB means the file is newer than the DB, not a data fault. Generated @@WHEN@@ from @@SOURCE@@.</small></header>
<main><div class="cards">@@CARDS@@</div>
<div class="tabs">@@TABS@@</div>
<p><input type=text id=q placeholder="filter project number"> Issue: @@CHIPS@@</p>
<table><thead><tr><th>Project</th><th>State</th><th>Verdict</th><th>Top causes</th><th>Trusses</th></tr></thead><tbody id=tb>@@ROWS@@</tbody></table>
</main>
<script>let t='Inconsistent';const cs=new Set();function f(){const q=document.getElementById('q').value.toLowerCase();document.querySelectorAll('#tb tr').forEach(r=>{const rc=(r.dataset.c||'').split('|');r.style.display=(r.dataset.s==t&&r.textContent.toLowerCase().indexOf(q)>=0&&(!cs.size||[...cs].some(c=>rc.includes(c))))?'':'none'})}
document.querySelectorAll('.tab').forEach(e=>e.onclick=()=>{t=e.dataset.t;document.querySelectorAll('.tab').forEach(x=>x.classList.toggle('on',x==e));f()});document.querySelectorAll('.chip').forEach(e=>e.onclick=()=>{const c=e.dataset.c;cs.has(c)?cs.delete(c):cs.add(c);e.classList.toggle('on');f()});document.getElementById('q').oninput=f;</script>
'@

if($Server -ieq 'SST'){$Server='.\SST'}
$cut=ParseVer $LegacyCutoffVersion
$root=(Get-Item -LiteralPath $ProjectsRoot).FullName.TrimEnd('\')
$stamp=(Get-Date).ToString('yyyyMMdd-HHmmss')
$outDir=$OutputDirectory.TrimEnd('\')+'-'+$stamp
[void](New-Item -ItemType Directory -Path (Join-Path $outDir 'projects') -Force)
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
$script:Conn.Open()

$work=@()
foreach($d in (Get-ChildItem -LiteralPath $root -Directory | Sort-Object Name)){
 if($d.Name -ieq 'DeletedProjects'){continue}
 if($ProjectId -and ($ProjectId -notcontains $d.Name)){continue}
 $arc=@(Get-ChildItem -LiteralPath $d.FullName -Filter '*.csdproj' -File | Where-Object {$_.Name -notmatch '^Attachments?_'} | Sort-Object Name)
 if($arc.Count -eq 0){continue}
 $work+=,@($d.Name,$arc[$arc.Count-1].FullName)
}
if($Limit -gt 0){$work=@($work | Select-Object -First $Limit)}
Write-Host ('Projects with archives: '+$work.Count)

$results=[Collections.Generic.List[object]]::new()
$used=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$i=0
try{
foreach($w in $work){
 $i++
 Write-Progress -Activity 'Analyzing projects' -Status ($w[0]+' ('+$i+' of '+$work.Count+')') -PercentComplete (100*$i/$work.Count)
 $number=$w[0]
 try{$o=AnalyzeProject $number $w[1] $cut}
 catch{Write-Warning ($number+": "+$_.Exception.Message+" @ "+$_.InvocationInfo.ScriptLineNumber);continue}
 $st=StateOf $o
 $o.state=$st
 $folder=SafeName $number;$n=2
 while(-not $used.Add($folder)){$folder=(SafeName $number)+'_'+$n;$n++}
 $o.folder=$folder
 $dir=Join-Path $outDir ('projects\'+$folder)
 [void](New-Item -ItemType Directory -Path (Join-Path $dir 'trusses') -Force)
 $byRule=[ordered]@{};$bySev=[ordered]@{}
 foreach($f in $o.findings){
  if($byRule.Contains($f.rule)){$byRule[$f.rule]++}else{$byRule[$f.rule]=1}
  if($bySev.Contains($f.sev)){$bySev[$f.sev]++}else{$bySev[$f.sev]=1}
 }
 $o.byRule=$byRule
 $summary=[ordered]@{project=$number;state=$st;archive=$o.archive;archiveTimestamp=$o.archiveTs;dbFilesCheckInId=$o.dbFilesCheckInId;trusses=$o.counts.trusses;verdict=(VerdictOf $o $st);trussesWithRealMismatch=$o.counts.bad;trussesStaleDb=$o.counts.stale;pre_r7_db_only_trusses=$o.counts.pre;findingsByRule=$byRule;findingsBySeverity=$bySev}
 WriteText (Join-Path $dir 'summary.json') (ToJson $summary 4)
 $flat=[Collections.Generic.List[object]]::new()
 foreach($f in $o.findings){$flat.Add([ordered]@{ruleId=$f.rule;ruleName=$Rules[$f.rule][0];rule=$Rules[$f.rule][1];project=$number;truss=$f.truss;field=$f.field;dbValue=$f.db;fileValue=$f.file;severity=$f.sev})}
 WriteText (Join-Path $dir 'findings.json') (ToJson $flat.ToArray() 4)
 $blobRows=[Collections.Generic.List[object]]::new()
 $tn=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
 foreach($t in $o.trusses){
  $tf=[Collections.Generic.List[object]]::new()
  $fl=[Collections.Generic.List[object]]::new()
  foreach($f in $t.findings){
   $tf.Add([ordered]@{ruleId=$f.rule;ruleName=$Rules[$f.rule][0];field=$f.field;db=$f.db;file=$f.file;severity=$f.sev;message=$f.msg})
   $fl.Add([object[]]@($Rules[$f.rule][0],$f.field,$f.db,$f.file,$f.sev,$f.msg,$f.group))
  }
  $tj=[ordered]@{truss=$t.name;project=$number;version=$t.version;cells=$t.cells;pre_r7_db_only=[bool]$t.pre;findings=$tf.ToArray();fields=$t.fields.ToArray()}
  $tfile=SafeName $t.name;$k=2
  while(-not $tn.Add($tfile)){$tfile=(SafeName $t.name)+'_'+$k;$k++}
  WriteText (Join-Path $dir ('trusses\'+$tfile+'.json')) (ToJson $tj 6)
  $cellArr=[string[]]@($Groups | ForEach-Object {$t.cells[$_]})
  $ver=if($t.version){$t.version}else{@{}}
  $blobRows.Add([object[]]@($t.name,$cellArr,$fl,$ver,$t.fields))
 }
 $sorted=@($blobRows | Sort-Object @{Expression={if($_[2].Count -gt 0){0}else{1}}},@{Expression={$_[0]}})
 $blob=if($sorted.Count -gt 0){(ToJson $sorted 8 -compress).Replace('</','<\/')}else{'[]'}
 $cc=[ordered]@{}
 foreach($f in $o.findings){$nm=$Rules[$f.rule][0];if($cc.Contains($nm)){$cc[$nm]++}else{$cc[$nm]=1}}
 $o.cats=$cc
 $chips=(($cc.Keys | Sort-Object | ForEach-Object {
  '<span class="chip" data-r="'+(Esc $_)+'">'+(Esc $_)+' ('+$cc[$_]+')</span>'
 }) -join '')
 if(-not $chips){$chips='no findings'}
 $headers=(($Groups | Where-Object {$_ -ne 'Version'} | ForEach-Object {'<th class=cell>'+$_+'</th>'}) -join '')
 $meta='archive '+(Esc $o.archive)+' | DB FilesCheckInId '+(Esc $o.dbFilesCheckInId)+' | '+$o.counts.trusses+' trusses | '+$o.counts.bad+' real mismatch | '+$o.counts.stale+' stale-DB | '+$o.counts.pre+' pre-r7 DB-only (not counted)'
 $page=$ProjectTemplate.Replace('@@CSS@@',$Css).Replace('@@TITLE@@',(Esc $number)).Replace('@@BADGE@@',$Badge[$st]).Replace('@@STATE@@',$st).Replace('@@META@@',$meta).Replace('@@CHIPS@@',$chips).Replace('@@HEADERS@@',$headers).Replace('@@GROUPS@@',(ToJson $Groups 2 -compress)).Replace('@@DATA@@',$blob)
 WriteText (Join-Path $dir 'project.html') $page
 $results.Add($o)
 $o.trusses=$null
}
if(-not $ProjectId -and $Limit -le 0){
 $have=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
 foreach($w in $work){[void]$have.Add($w[0])}
 $dbp=Q 'SELECT p.ProjectNumber,p.FilesCheckInId,(SELECT COUNT(*) FROM dbo.ComponentHeader h WHERE h.ProjectEventKey=p.EventKey) AS Trusses FROM dbo.Project p ORDER BY p.ProjectNumber' @{}
 foreach($r in $dbp.Rows){
  $pn=[string]$r['ProjectNumber']
  if($have.Contains($pn)){continue}
  $ci=if($r['FilesCheckInId'] -is [DBNull]){$null}else{[string]$r['FilesCheckInId']}
  $results.Add(@{project=$pn;state='DB only';folder=$null;dbFilesCheckInId=$ci;counts=@{trusses=[int]$r['Trusses'];bad=0;stale=0;pre=0};cats=[ordered]@{}})
 }
}
}finally{$script:Conn.Dispose();Write-Progress -Activity 'Analyzing projects' -Completed}

$order=@{'Inconsistent'=0;'Stale DB'=1;'Not in DB'=2;'Consistent'=3;'Pre-r7'=4;'DB only'=5}
$ranked=@($results | Sort-Object @{Expression={$order[$_.state]}},@{Expression={-($_.counts.bad+$_.counts.stale)}},@{Expression={$_.project}})
$cnt=@{};foreach($s in $order.Keys){$cnt[$s]=0};foreach($o in $ranked){$cnt[$o.state]++}
$tabNames=@('Inconsistent','Stale DB','Not in DB','Consistent','Pre-r7','DB only')
$tabCount=$cnt
$cards=(($tabNames | ForEach-Object {'<div class=card><b>'+$tabCount[$_]+'</b>'+$_+'</div>'}) -join '')
$tabs=(($tabNames | ForEach-Object {'<div class="tab'+$(if($_ -eq 'Inconsistent'){' on'}else{''})+'" data-t="'+$_+'">'+$_+' ('+$tabCount[$_]+')</div>'}) -join '')
$allChips=(($Rules.Values | ForEach-Object {'<span class="chip" data-c="'+(Esc $_[0])+'" title="'+(Esc $_[1])+'">'+(Esc $_[0])+'</span>'}) -join '')
$sb=[Text.StringBuilder]::new()
foreach($o in $ranked){
 $top=(($o.cats.GetEnumerator() | Sort-Object @{Expression={-$_.Value}},@{Expression={$_.Key}} | Select-Object -First 2 | ForEach-Object {'<span class="badge b-pre">'+(Esc $_.Key)+' '+$_.Value+'</span>'}) -join '')
 [void]$sb.Append('<tr data-s="'+$o.state+'" data-c="'+(Esc ($o.cats.Keys -join '|'))+'"><td>'+$(if($o.folder){'<a href="projects/'+(Esc $o.folder)+'/project.html">'+(Esc $o.project)+'</a>'}else{Esc $o.project})+'</td><td><span class="badge '+$Badge[$o.state]+'">'+$o.state+'</span></td><td>'+(Esc (VerdictOf $o $o.state))+'</td><td>'+$top+'</td><td>'+$o.counts.trusses+'</td></tr>')
}
$index=$IndexTemplate.Replace('@@CSS@@',$Css).Replace('@@COUNT@@',[string]$ranked.Count).Replace('@@WHEN@@',(Get-Date).ToString('yyyy-MM-dd HH:mm')).Replace('@@SOURCE@@',(Esc $root)).Replace('@@CARDS@@',$cards).Replace('@@TABS@@',$tabs).Replace('@@CHIPS@@',$allChips).Replace('@@ROWS@@',$sb.ToString())
WriteText (Join-Path $outDir 'index.html') $index
Write-Host ''
Write-Host ('Projects: '+$ranked.Count)
foreach($s in $order.Keys){Write-Host ('  '+$s+': '+$cnt[$s])}
$reportUri=([System.Uri]([IO.Path]::GetFullPath((Join-Path $outDir 'index.html')))).AbsoluteUri
$escape=[char]27
Write-Host ("Report: ${escape}]8;;${reportUri}${escape}\${reportUri}${escape}]8;;${escape}\")

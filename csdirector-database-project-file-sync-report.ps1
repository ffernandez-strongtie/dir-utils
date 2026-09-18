#requires -Version 5.1
<#
.SYNOPSIS
Read-only csdproj/database comparison with linked HTML reports.
.EXAMPLE
.\csdirector-database-project-file-sync-report.ps1
.EXAMPLE
.\csdirector-database-project-file-sync-report.ps1 -ProjectId S-1206,S-1328
.EXAMPLE
.\csdirector-database-project-file-sync-report.ps1 -Limit 2
.EXAMPLE
.\csdirector-database-project-file-sync-report.ps1 -Evidence True
.EXAMPLE
.\csdirector-database-project-file-sync-report.ps1 -Timestamp
.EXAMPLE
.\csdirector-database-project-file-sync-report.ps1 -Database CustomerDatabase
.NOTES
Output defaults to output-TIMESTAMP under the current working directory. A UTC timestamp and unique suffix are appended to custom output paths too. JSON files are extracted from copied
ZIPs into temporary folders, removed by default. Use -Evidence True to retain copied ZIPs, extracted JSON and full comparison inventories. Originals and SQL records are never changed.
Timestamp comparisons only run with -Timestamp. Both timestamp interpretations default to the current machine timezone.
Use -FileTimeZone and -DatabaseTimeZone to override for data from another system.
#>
[CmdletBinding()]
param(
 [string]$ProjectsRoot='C:\SST\Server\Projects',
 [string]$Server='.\SST',
 [string]$Database='CSDatabaseServer',
 [string]$OutputDirectory=(Join-Path (Get-Location).Path 'output'),
 [string[]]$ProjectId,
 [ValidateSet('True','False')][string]$Evidence='False',
 [switch]$Timestamp,
 [ValidateRange(0,2147483647)][int]$Limit=0,
 [string]$FileTimeZone=([TimeZoneInfo]::Local.Id),
 [string]$DatabaseTimeZone=([TimeZoneInfo]::Local.Id),
 [ValidateRange(0,86400)][double]$ToleranceSeconds=120,
 [ValidateRange(0,1)][double]$NumericTolerance=0.000001
)
$ErrorActionPreference='Stop'
$retainEvidence=[bool]::Parse($Evidence)
Write-Host ''
Write-Host 'CSDirector: Database and Project file Sync Report' -ForegroundColor Cyan
Write-Host 'Searches project folders and compares archived trusses with SQL records using Windows authentication.'
Write-Host 'Checks names, selected truss fields, individual truss-file checksums, estimated board footage, suspected duplicate DB pieces.'
 if($Timestamp){Write-Host 'Timestamp comparison is on.'}else{Write-Host 'Timestamp comparison is off. Add -Timestamp to include it.'}
Write-Host 'Creates an HTML overview and a detail page per project in a timestamped output folder.'
Write-Host 'Original project files and database records are unchanged. DeletedProjects folders and Attachment_/Attachments_ archives are skipped.'
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
if($Server -ieq 'SST'){$Server='.\SST'}
Add-Type -AssemblyName System.IO.Compression.FileSystem
$map=[ordered]@{Quantity='ComponentQuantity';NumPlies='TrussPlyCount';Thickness='TrussThicknessInches';OverallTrussHeight='OverallTrussHeightInches';LeftOverhang='LeftOverhangInches';RightOverhang='RightOverhangInches';LeftCantilever='LeftCantileverInches';RightCantilever='RightCantileverInches';LeftHeelHeight='LeftHeelHeightInches';RightHeelHeight='RightHeelHeightInches';PitchLeftTopOver12='LeftTopPitchOver12';PitchRightTopOver12='RightTopPitchOver12';PitchLeftBottomOver12='LeftBottomPitchOver12';PitchRightBottomOver12='RightBottomPitchOver12';IsAttic='IsAttic';IsGable='IsGable';IsGirder='IsGirder';IsFlipped='IsFlipped';OCSpacing='OCSpacing'}
function Enc($v){[Net.WebUtility]::HtmlEncode([string]$v)}
function RawDate($v){if($null -eq $v -or $v -is [DBNull]){return ''};([datetime]$v).ToString('yyyy-MM-dd HH:mm:ss.fffffff')}
function UtcDate([datetime]$v,$zone){$wall=[datetime]::SpecifyKind($v,[DateTimeKind]::Unspecified);if($zone.IsInvalidTime($wall) -or $zone.IsAmbiguousTime($wall)){throw 'Ambiguous or invalid daylight-saving wall time'};[TimeZoneInfo]::ConvertTimeToUtc($wall,$zone)}
function SignedTime($seconds){if($null -eq $seconds){return 'Unavailable'};$span=[timespan]::FromSeconds([math]::Abs($seconds));$sign=if($seconds -lt 0){'-'}else{'+'};'{0}{1}d {2:00}h {3:00}m {4:00}s' -f $sign,$span.Days,$span.Hours,$span.Minutes,$span.Seconds}
function CellColor($column,$value){
 $text=[string]$value
 if($column -eq 'Status'){
  if($text -match 'errors|incomplete|Missing from csdproj'){return 'error'}
  if($text -match 'Mixed'){return 'mixed'}
  if($text -match 'Older'){return 'older'}
  if($text -match 'Newer'){return 'newer'}
  if($text -match 'Within tolerance'){return 'close'}
 }
 if($column -eq 'Only in database' -and [double]$value -gt 0){return 'error'}
 if($column -eq 'Older Timestamp' -and [double]$value -gt 0){return 'older'}
 if($column -eq 'Newer Timestamp' -and [double]$value -gt 0){return 'newer'}
 if($column -eq 'Within tolerance' -and [double]$value -gt 0){return 'close'}
 if($column -eq 'Unavailable' -and [double]$value -gt 0){return 'unknown'}
 if($column -in @('Minimum difference','Maximum difference','Difference')){
  if($text.StartsWith('-')){return 'older'}
  if($text.StartsWith('+')){return 'newer'}
  return 'unknown'
 }
 return ''
}
function HealthyProject($row){
 $paired=[int]$row.TdlMatchedNames
 return ($paired -gt 0 -and $row.TdlTrussCount -eq $paired -and $row.DbTrussCount -eq $paired -and
  $row.'Data synced' -eq $paired -and $row.ChecksumSame -eq $paired -and
  $row.'Data differs' -eq 0 -and $row.ChecksumDifferent -eq 0 -and $row.ChecksumMissing -eq 0 -and $row.ChecksumUnavailable -eq 0 -and
  $row.Ambiguous -eq 0 -and $row.Status -notmatch 'errors|incomplete|ambiguous|Project not in DB')
}
function BoardFeetStatus($row){
 if($null -eq $row.SameBasisBoardFeetDifference){return 'BDFT comparison unavailable'}
 if($row.SameBasisDifferentTrusses -gt 0 -or [math]::Abs([decimal]$row.SameBasisBoardFeetDifference) -gt 0.000001){return 'BDFT differs on file basis'}
 if($row.BoardFeetQuantityDifferences -gt 0){return 'BDFT quantities differ or are unavailable'}
 return 'BDFT agrees'
}

function StatusHtml($row){
 $parts=[Collections.Generic.List[string]]::new()
 if($row.Healthy){$parts.Add('<div class="close healthy-status" style="padding:5px;margin:3px 0"><b>Healthy</b></div>')}elseif(HealthyProject $row){$parts.Add('<div class="close" style="padding:5px;margin:3px 0"><b>Inventory / fields / files match</b></div>')}
 if($row.BoardFeetSettingsDifferent -eq $true){
  $parts.Add('<div class="notice" style="padding:5px;margin:3px 0">BDFT: file and Director length methods differ. Compare csdproj with DB (calculated).</div>')
 }elseif($null -ne $row.BoardFeetDifference -and [math]::Abs([decimal]$row.BoardFeetDifference) -gt 0.000001){
  $parts.Add('<div class="notice" style="padding:5px;margin:3px 0">BDFT: Director stored total differs. Compare csdproj with DB (calculated).</div>')
 }
 if($row.Status -match 'errors|incomplete|Project not in DB'){$parts.Add('<div class="error" style="padding:5px;margin:3px 0">'+(Enc $row.Status)+'</div>')}
 if($row.BoardFeetStatus -and -not $row.Healthy){$parts.Add('<div class="'+$(if($row.BoardFeetStatus -eq 'BDFT agrees'){'close'}else{'notice'})+'" style="padding:5px;margin:3px 0">'+(Enc $row.BoardFeetStatus)+'</div>')}
 if($row.'Only in database' -gt 0){$parts.Add('<div class="error" style="padding:5px;margin:3px 0">Missing from csdproj: '+$row.'Only in database'+' DB truss names</div>')}
 if($Timestamp){
 $mixed=$row.'Older Timestamp' -gt 0 -and $row.'Newer Timestamp' -gt 0
 if($mixed){$parts.Add('<div class="mixed" style="padding:5px;margin:3px 0">Mixed timestamps</div>')}
 foreach($pair in @(@('Older Timestamp','older'),@('Newer Timestamp','newer'),@('Within tolerance','close'),@('Unavailable','unknown'))){
  if($mixed -and $pair[1] -in @('older','newer')){continue}
  if($row.($pair[0]) -gt 0){$parts.Add('<div class="'+$pair[1]+'" style="padding:5px;margin:3px 0">'+$pair[0]+': '+$row.($pair[0])+'</div>')}
 }
 }
 if(-not $parts.Count){if($row.Status -eq 'Healthy'){return 'Needs review'};return Enc $row.Status}
 $parts -join ''
}
function ReadArchiveXml($entry){
 $settings=[Xml.XmlReaderSettings]::new();$settings.DtdProcessing=[Xml.DtdProcessing]::Prohibit;$settings.XmlResolver=$null
 $stream=$entry.Open();$reader=$null
 try{$reader=[Xml.XmlReader]::Create($stream,$settings);$doc=[Xml.XmlDocument]::new();$doc.XmlResolver=$null;$doc.Load($reader);return ,$doc}
 finally{if($reader){$reader.Dispose()};$stream.Dispose()}
}
function FormulaBoardFeet([decimal]$thickness,[decimal]$width,[decimal]$length){
 # dbo.CalculateBoardFeet takes decimal(28,8) inputs. Its products and result
 # are decimal(38,6). Products round away from zero; the observed SQL decimal
 # division truncates to six places. Direct SQL comparison tests cover both.
 $round=[MidpointRounding]::AwayFromZero
 $thickness=[math]::Round($thickness,8,$round);$width=[math]::Round($width,8,$round);$length=[math]::Round($length,8,$round)
 $area=[math]::Round($thickness*$width,6,$round)
 $volume=[math]::Round($area*$length,6,$round)
 [math]::Truncate(($volume/144d)*1000000d)/1000000d
}
function ArchiveBoardFeetBasis($bdft){
 if($null -eq $bdft.Total){return 'Unavailable'}
 $modes=@($bdft.Trusses | Select-Object -ExpandProperty LengthMode -Unique)
 if($modes.Count -eq 1){return $modes[0]}
 if($modes.Count -gt 1){return 'Mixed'}
 return 'Unverified'
}
function FormatBoardFeet($value,[string]$basis,[int]$places=2){
 if($null -eq $value){return [string][char]0x2014}
 $number=([decimal]$value).ToString(('N'+$places),[Globalization.CultureInfo]::InvariantCulture)
 switch($basis){
  'Pick Length' {return $number+'*'}
  'Actual Length' {return $number}
  'Mixed' {return $number+' (Mixed)'}
  'Either Length' {return $number}
  default {return $number+' (basis unverified)'}
 }
}
function DatabaseBoardFeetBasis($pieceTable,$databaseRows,$cached,[string]$preset){
 # Verify every project piece cache against its stored length-specific caches.
 # A current preset alone cannot establish a cached total's basis.
 $result=[ordered]@{PieceBasis='Unverified';CachedBasis='Unverified';Note='DB BDFT basis could not be verified.'}
 try{
  if($null -eq $pieceTable -or $pieceTable.Rows.Count -eq 0){throw 'No DB pieces available'}
  $byComponent=@{};$allPick=$true;$allActual=$true;$hasPick=$false;$hasActual=$false
  foreach($piece in $pieceTable.Rows){
   foreach($field in @('CalculatedBoardFeet','CalculatedPickLengthBoardFeet','CalculatedOverallLengthBoardFeet','PlyCount')){
    if($piece[$field] -is [DBNull]){throw ('Missing DB piece cache: '+$field)}
   }
   $isPick=[decimal]$piece.CalculatedBoardFeet -eq [decimal]$piece.CalculatedPickLengthBoardFeet
   $isActual=[decimal]$piece.CalculatedBoardFeet -eq [decimal]$piece.CalculatedOverallLengthBoardFeet
   if(-not $isPick -and -not $isActual){throw 'A DB piece cache matches neither length-specific cache'}
   $allPick=$allPick -and $isPick;$allActual=$allActual -and $isActual
   $hasPick=$hasPick -or ($isPick -and -not $isActual);$hasActual=$hasActual -or ($isActual -and -not $isPick)
   $key=[string]$piece.ComponentKey
   $byComponent[$key]+=[decimal]$piece.CalculatedBoardFeet*[decimal]$piece.PlyCount
  }
  $basis=if($allPick -and -not $allActual){'Pick Length'}elseif($allActual -and -not $allPick){'Actual Length'}elseif($hasPick -and $hasActual){'Mixed'}else{'Either Length'}
  $result.PieceBasis=$basis
  $seen=@{};$projectSum=0d
  foreach($db in $databaseRows){
   $key=[string]$db.ComponentHeaderKey
   if($seen.ContainsKey($key) -or $db.ComponentQuantity -is [DBNull] -or [decimal]$db.ComponentQuantity -lt 0 -or [decimal]$db.ComponentQuantity -ne [math]::Truncate([decimal]$db.ComponentQuantity)){throw 'DB component quantities unavailable, invalid or ambiguous'}
   $seen[$key]=$true
   if(-not $byComponent.ContainsKey($key)){throw 'A DB component has no verified piece sum'}
   $projectSum+=[decimal]$byComponent[$key]*[decimal]$db.ComponentQuantity
  }
  foreach($key in $byComponent.Keys){if(-not $seen.ContainsKey($key)){throw 'DB piece component missing from project inventory'}}
  if($null -ne $cached -and [math]::Abs([double]$cached-$projectSum) -le 0.000001){
   $result.CachedBasis=$basis
   $result.Note='DB stored total reconciles with current piece caches, piece plies and DB component quantities (tolerance 0.000001 BDFT). The marker describes agreement with stored length-specific caches, not historical provenance or a fresh cache validation. Either length means both stored length-specific caches agree for every piece.'
  }else{$result.Note='DB piece basis verified, but the stored project total does not reconcile with current DB quantities; stored-total basis remains unverified.'}
 }catch{$result.Note+=' '+$_.Exception.Message}
 [pscustomobject]$result
}

function EstimateBoardFeet($zip){
 # Formula-based geometry calculation, with archive material and mode evidence.
 # Project membership still comes from layouts, not every retained JSON.
 $quantities=@{};$trussRows=[Collections.Generic.List[object]]::new();$pieceRows=[Collections.Generic.List[object]]::new()
 try{
  $studioEntries=@($zip.Entries | Where-Object {$_.FullName -ieq 'Studio.json'})
  if($studioEntries.Count -ne 1){throw 'Studio.json layout membership unavailable or duplicated'}
  $reader=[IO.StreamReader]::new($studioEntries[0].Open())
  try{$studio=$reader.ReadToEnd() | ConvertFrom-Json}finally{$reader.Dispose()}
  if($null -eq $studio.PSObject.Properties['Layouts']){throw 'Layout membership missing'}
  foreach($layout in $studio.Layouts){
   if($null -eq $layout.PSObject.Properties['Trusses']){throw 'Layout truss list missing'}
   foreach($truss in $layout.Trusses){
    if([string]::IsNullOrWhiteSpace($truss.Name) -or $null -eq $truss.Qty -or [decimal]$truss.Qty -lt 0 -or [decimal]$truss.Qty -ne [math]::Truncate([decimal]$truss.Qty)){throw 'Invalid layout quantity'}
    $quantities[$truss.Name]+=[decimal]$truss.Qty
   }
  }
  $entries=@{}
  foreach($entry in $zip.Entries){if($entry.FullName -match '(?i)(^|/)Trusses/generated/[^/]+\.json$'){
   $name=[IO.Path]::GetFileNameWithoutExtension($entry.Name)
   $entries[$name]=@($entries[$name] | Where-Object {$null -ne $_})+@($entry)
  }}
  $inventoryEntries=@($zip.Entries | Where-Object FullName -IEQ 'Presets/TrussStudio/LumberInventory.xml')
  if($inventoryEntries.Count -ne 1){throw 'LumberInventory.xml missing or duplicated; material dimensions cannot be verified'}
  $inventory=ReadArchiveXml $inventoryEntries[0];$materials=@{}
  foreach($material in $inventory.SelectNodes('//LumberMaterialList/Lumber')){
   $key=([guid]$material.GetAttribute('Key')).ToString('N')
   if($materials.ContainsKey($key)){throw ('Duplicate lumber material GUID '+$key)}
   $materials[$key]=$material
  }
  $tdl=@{};foreach($entry in $zip.Entries){if($entry.FullName -match '(?i)^Trusses/[^/]+\.tdlTruss$'){$n=[IO.Path]::GetFileNameWithoutExtension($entry.Name);$tdl[$n]=@($tdl[$n] | Where-Object {$null -ne $_})+@($entry)}}
  $total=0d;$envMode=$null;$actual=0;$pick=0;$fallback=0
  foreach($name in @($quantities.Keys | Sort-Object)){
   if($quantities[$name] -eq 0){continue}
   $matches=@($entries[$name] | Where-Object {$null -ne $_})
   if($matches.Count -ne 1){throw ('Missing or duplicate generated JSON for layout truss '+$name)}
   $reader=[IO.StreamReader]::new($matches[0].Open())
   try{$json=$reader.ReadToEnd() | ConvertFrom-Json}finally{$reader.Dispose()}
   if($null -eq $json.PSObject.Properties['PieceData'] -or $null -eq $json.PieceData -or @($json.PieceData).Count -eq 0){throw ('Missing or empty PieceData: '+$name)}
   $trussEntries=@($tdl[$name] | Where-Object {$null -ne $_});if($trussEntries.Count -ne 1){throw ('Missing or duplicate canonical tdlTruss for layout truss '+$name)}
   $trussXml=ReadArchiveXml $trussEntries[0];$modeNodes=$trussXml.SelectNodes('//OtheruseExactBdft');$modeSource=$trussEntries[0].FullName
   if($modeNodes.Count -gt 1){throw ('Duplicate OtheruseExactBdft setting: '+$name)}
   if($modeNodes.Count -eq 1){$modeValue=$modeNodes[0].GetAttribute('Value')}
   else{
    if($null -eq $envMode){
     $envEntries=@($zip.Entries | Where-Object FullName -IEQ 'Presets/TrussStudio/EnvData.tdlEnv')
     if($envEntries.Count -ne 1){throw ('No saved board-footage mode for '+$name)}
     $envXml=ReadArchiveXml $envEntries[0];$envNodes=$envXml.SelectNodes('//useExactBdft')
     if($envNodes.Count -ne 1){throw 'Environment useExactBdft setting missing or duplicated'}
     $envMode=$envNodes[0].GetAttribute('Value')
    }
    $modeValue=$envMode;$modeSource='Presets/TrussStudio/EnvData.tdlEnv (fallback)';$fallback++
   }
   $lengthField=switch($modeValue.ToLowerInvariant()){'true'{'Overall'} 'false'{'PickLength'} default{throw ('Invalid board-footage mode: '+$name)}}
   $mode=if($lengthField -eq 'Overall'){$actual++;'Actual Length'}else{$pick++;'Pick Length'}
   $trussTotal=0d;$pieceIndex=0
   foreach($piece in $json.PieceData){
    $pieceIndex++;$key=([guid]$piece.GoldenGuid).ToString('N')
    if(-not $materials.ContainsKey($key)){throw ('Lumber material absent from inventory: '+$name+' / '+$piece.EngineeringLabel+' / '+$piece.GoldenGuid)}
    $material=$materials[$key];$sclValue=$material.GetAttribute('SCL').ToLowerInvariant()
    if($sclValue -eq 'true'){
     $thickness=[decimal]::Parse($material.GetAttribute('Thickness'),[Globalization.CultureInfo]::InvariantCulture)
     $width=[decimal]::Parse($material.GetAttribute('Width'),[Globalization.CultureInfo]::InvariantCulture);$dimensionBasis='Actual engineered-lumber dimensions'
    }elseif($sclValue -eq 'false'){
     $size=[regex]::Match($material.GetAttribute('Size'),'(?i)^\s*(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)\s*$')
     if(-not $size.Success){throw ('Invalid nominal lumber size: '+$name+' / '+$material.GetAttribute('Size'))}
     $thickness=[decimal]::Parse($size.Groups[1].Value,[Globalization.CultureInfo]::InvariantCulture)
     $width=[decimal]::Parse($size.Groups[2].Value,[Globalization.CultureInfo]::InvariantCulture);$dimensionBasis='Nominal dimensional-lumber size'
    }else{throw ('Missing or invalid SCL material flag: '+$name)}
    if($thickness -le 0 -or $width -le 0 -or $null -eq $piece.Lengths.$lengthField -or [string]::IsNullOrWhiteSpace([string]$piece.Lengths.$lengthField) -or $piece.Lengths.$lengthField -is [bool] -or $null -eq $piece.PlyCount -or $piece.PlyCount -is [bool]){throw ('Missing/invalid dimensions, required '+$lengthField+' length or plies: '+$name+' / '+$piece.EngineeringLabel+'; no alternate length substituted')}
    $length=[decimal]$piece.Lengths.$lengthField;$plies=[decimal]$piece.PlyCount
    if($length -lt 0 -or $plies -lt 1 -or $plies -ne [math]::Truncate($plies)){throw ('Invalid '+$lengthField+' length/plies: '+$name+' / '+$piece.EngineeringLabel)}
    $onePly=FormulaBoardFeet $thickness $width $length;$allPlies=$onePly*$plies;$trussTotal+=$allPlies
    $pieceRows.Add([pscustomobject]@{Truss=$name;PieceIndex=$pieceIndex;EngineeringLabel=$piece.EngineeringLabel;MaterialGuid=$piece.GoldenGuid;MaterialName=$material.GetAttribute('Name');DimensionBasis=$dimensionBasis;Thickness=$thickness;Width=$width;LengthMode=$mode;LengthInches=$length;PlyCount=$plies;OnePlyBoardFeet=$onePly;AllPliesBoardFeet=$allPlies;LayoutQuantity=$quantities[$name];ProjectBoardFeet=$allPlies*$quantities[$name]})
   }
   $total+=$trussTotal*$quantities[$name]
   $trussRows.Add([pscustomobject]@{Truss=$name;LengthMode=$mode;ModeSource=$modeSource;PieceCount=$pieceIndex;LayoutQuantity=$quantities[$name];OneTrussBoardFeet=$trussTotal;ProjectBoardFeet=$trussTotal*$quantities[$name]})
  }
  [pscustomobject]@{Total=$total;Quantities=$quantities;Trusses=$trussRows.ToArray();Pieces=$pieceRows.ToArray();Note=('Formula-based estimate: archive material dimensions and saved length settings; SQL decimal arithmetic; piece plies; summed Studio.json layout quantities. '+$actual+' actual-length trusses, '+$pick+' pick-length trusses; '+$fallback+' environment fallbacks. This is not a stored Studio total; Director pricing exceptions and project-cache aggregation are not reproduced.')}
 }catch{[pscustomobject]@{Total=$null;Quantities=$quantities;Trusses=$trussRows.ToArray();Pieces=$pieceRows.ToArray();Note=('Formula-based estimate unavailable: '+$_.Exception.Message+'; partial totals withheld')}}
}
function ReconcileBoardFeet($bdft,$pieceTable,$duplicates,$databaseRows,$cached,$databaseMode){
 $result=[ordered]@{DatabaseLengthMode=$databaseMode;DatabasePieceSum=$null;DuplicateContribution=$null;DatabaseWithoutDuplicates=$null;CachedLessDuplicates=$null;Trusses=@();Note=''}
 try{
  if($null -eq $bdft.Total){throw 'Complete archive formula calculation unavailable'}
  if($null -eq $pieceTable -or $null -eq $duplicates.AffectedTrusses){throw 'DB piece data or duplicate detection unavailable'}
  $headers=@{};foreach($db in $databaseRows){$headers[$db.Name]=@($headers[$db.Name] | Where-Object {$null -ne $_})+@($db)}
  $pieces=@{};foreach($piece in $pieceTable.Rows){$key=[string]$piece.ComponentKey;$pieces[$key]=@($pieces[$key] | Where-Object {$null -ne $_})+@($piece)}
  $extra=@{};foreach($group in $duplicates.Groups){
   $key=[string]$group.ComponentKey
   $extra[$key]+=[decimal]$group.MatchingData.CalculatedBoardFeet*[decimal]$group.MatchingData.PlyCount*[decimal]$group.ExtraPieceRows
  }
  $before=0d;$excess=0d;$rows=[Collections.Generic.List[object]]::new()
  foreach($source in $bdft.Trusses){
   $matches=@($headers[$source.Truss] | Where-Object {$null -ne $_})
   if($matches.Count -ne 1 -or $matches[0].ComponentKey -is [DBNull]){throw ('No unique DB truss counterpart for layout member '+$source.Truss)}
   $componentKey=[string]$matches[0].ComponentHeaderKey;$oneTruss=0d
   foreach($piece in @($pieces[$componentKey] | Where-Object {$null -ne $_})){
    if($piece.CalculatedBoardFeet -is [DBNull] -or $piece.PlyCount -is [DBNull]){throw ('Missing cached piece footage or plies: '+$source.Truss)}
    $oneTruss+=[decimal]$piece.CalculatedBoardFeet*[decimal]$piece.PlyCount
   }
   $duplicateOneTruss=[decimal]$extra[$componentKey];$quantity=[decimal]$source.LayoutQuantity
   $before+=$oneTruss*$quantity;$excess+=$duplicateOneTruss*$quantity
   $rows.Add([pscustomobject]@{Truss=$source.Truss;LayoutQuantity=$quantity;CsdprojMode=$source.LengthMode;CsdprojFormula=$source.ProjectBoardFeet;DbPieceSum=$oneTruss*$quantity;DuplicateContribution=$duplicateOneTruss*$quantity;DbWithoutDuplicates=($oneTruss-$duplicateOneTruss)*$quantity;DifferenceAfterDeduplication=$source.ProjectBoardFeet-($oneTruss-$duplicateOneTruss)*$quantity})
  }
  $result.DatabasePieceSum=$before;$result.DuplicateContribution=$excess;$result.DatabaseWithoutDuplicates=$before-$excess
  if($null -ne $cached){$result.CachedLessDuplicates=[decimal]$cached-$excess}
  $result.Trusses=$rows.ToArray()
  $result.Note='Both piece sums use the same Studio.json layout quantities and membership as the csdproj formula calculation. DB sums use cached Piece.CalculatedBoardFeet multiplied by Piece.PlyCount; no refresh occurs. Duplicate contribution excludes repetitions beyond the first exact business-data match. Cached less duplicates is a diagnostic subtraction only: the original project-cache quantity aggregation and provenance are not verified. It is not a predicted application refresh result.'
 }catch{$result.Note='BDFT reconciliation unavailable: '+$_.Exception.Message+'; partial totals withheld.'}
 [pscustomobject]$result
}
function CompareBoardFeetOnFileBasis($bdft,$pieceTable,$databaseRows,$databaseMode){
 $result=[ordered]@{DatabaseTotal=$null;Difference=$null;DifferentTrusses=$null;QuantityDifferences=$null;SettingsDifferent=$null;Trusses=@();Note=''}
 try{
  if($null -eq $bdft.Total -or $null -eq $pieceTable){throw 'Complete archive calculation and DB pieces required'}
  $headers=@{};foreach($db in $databaseRows){$headers[$db.Name]=@($headers[$db.Name] | Where-Object {$null -ne $_})+@($db)}
  $pieces=@{};foreach($piece in $pieceTable.Rows){$key=[string]$piece.ComponentKey;$pieces[$key]=@($pieces[$key] | Where-Object {$null -ne $_})+@($piece)}
  $total=0d;$different=0;$quantityDifferences=0;$settingsDifferent=$false;$rows=[Collections.Generic.List[object]]::new()
  foreach($source in $bdft.Trusses){
   $matches=@($headers[$source.Truss] | Where-Object {$null -ne $_})
   if($matches.Count -ne 1 -or $matches[0].ComponentKey -is [DBNull]){throw ('Unique DB truss counterpart required: '+$source.Truss)}
   if($source.LengthMode -notin @('Pick Length','Actual Length')){throw ('Invalid file length mode: '+$source.Truss)}
   $lengthField=if($source.LengthMode -eq 'Pick Length'){'PickLengthInches'}else{'OverallLengthInches'}
   $members=@($pieces[[string]$matches[0].ComponentHeaderKey] | Where-Object {$null -ne $_})
   if(-not $members.Count){throw ('No DB pieces for file layout member: '+$source.Truss)}
   $oneTruss=0d
   foreach($piece in $members){
    foreach($field in @($lengthField,'PlyCount','BdftThickness','BdftWidth')){if($piece[$field] -is [DBNull]){throw ('Missing DB '+$field+': '+$source.Truss+' / '+$piece.EngineeringLabel)}}
    if([decimal]$piece[$lengthField] -lt 0 -or [decimal]$piece.PlyCount -lt 1 -or [decimal]$piece.PlyCount -ne [math]::Truncate([decimal]$piece.PlyCount) -or [decimal]$piece.BdftThickness -le 0 -or [decimal]$piece.BdftWidth -le 0){throw ('Invalid DB dimensions, length or plies: '+$source.Truss)}
    $oneTruss+=(FormulaBoardFeet ([decimal]$piece.BdftThickness) ([decimal]$piece.BdftWidth) ([decimal]$piece[$lengthField]))*[decimal]$piece.PlyCount
   }
   $contribution=$oneTruss*[decimal]$source.LayoutQuantity;$total+=$contribution
   $difference=$source.ProjectBoardFeet-$contribution
   if([math]::Abs($difference) -gt 0.000001){$different++}
   $dbQuantity=$matches[0].ComponentQuantity
   if($dbQuantity -is [DBNull] -or [decimal]$dbQuantity -ne [decimal]$source.LayoutQuantity){$quantityDifferences++}
   if(($databaseMode -eq 'Actual Length' -and $source.LengthMode -ne 'Actual Length') -or ($databaseMode -like 'Pick Length*' -and $source.LengthMode -ne 'Pick Length')){$settingsDifferent=$true}
   $rows.Add([pscustomobject]@{Truss=$source.Truss;LengthMode=$source.LengthMode;LayoutQuantity=$source.LayoutQuantity;DatabaseQuantity=if($dbQuantity -is [DBNull]){$null}else{$dbQuantity};FilePieceCount=$source.PieceCount;DatabasePieceCount=$members.Count;FileBoardFeet=$source.ProjectBoardFeet;DatabaseOnFileBasis=$contribution;Difference=$difference})
  }
  $result.DatabaseTotal=$total;$result.Difference=[decimal]$bdft.Total-[decimal]$total;$result.Trusses=$rows.ToArray();$result.DifferentTrusses=$different;$result.QuantityDifferences=$quantityDifferences
  $result.SettingsDifferent=if($databaseMode -eq 'Actual Length' -or $databaseMode -like 'Pick Length*'){$settingsDifferent}else{$null}
  $result.Note='Both sources use the dbo.CalculateBoardFeet geometry formula and SQL decimal rounding, each archive truss length mode, piece plies and the same Studio.json layout quantities. DB dimensions come from its lumber material records. Director pricing exceptions are not applied to either geometry total. Duplicate DB pieces remain included. DB-only trusses are excluded from this layout-based check and reported separately; quantity differences are shown explicitly. Agreement is checked per truss as well as in total (0.000001 BDFT tolerance). This is a geometry comparison, not a prediction of a Director cache refresh.'
 }catch{$result.Note='Same-basis comparison unavailable: '+$_.Exception.Message+'; no alternate length substituted and partial totals withheld.'}
 [pscustomobject]$result
}

function BoardFeetDetails($bdft,$reconciliation,$cached,$basisEvidence,$fileBasisComparison){
 $items=@(
  [pscustomobject]@{Calculation='csdproj formula-based estimate';BoardFeet=$bdft.Total;Meaning='Archive material dimensions, saved truss length modes, SQL rounding, piece plies and layout quantities; not a stored Studio total.'}
  [pscustomobject]@{Calculation='DB recalculated on file length basis';BoardFeet=$fileBasisComparison.DatabaseTotal;Meaning='Same geometry formula and SQL rounding as the file; DB material dimensions, file length modes and layout quantities. Pricing exceptions excluded on both sides.'}
  [pscustomobject]@{Calculation='DB stored project total';BoardFeet=$cached;Meaning='Project.CachedBoardFeet; original aggregate calculation not reproduced.'}
  [pscustomobject]@{Calculation='DB piece-cache sum';BoardFeet=$reconciliation.DatabasePieceSum;Meaning='SUM(CalculatedBoardFeet x PlyCount x the same layout quantities).'}
  [pscustomobject]@{Calculation='Suspected duplicate contribution';BoardFeet=$reconciliation.DuplicateContribution;Meaning='Footage from extra exact-matching piece rows, using the same layout quantities.'}
  [pscustomobject]@{Calculation='DB piece-cache sum without duplicates';BoardFeet=$reconciliation.DatabaseWithoutDuplicates;Meaning='Keep one row per suspected duplicate group; calculated in memory.'}
  [pscustomobject]@{Calculation='DB stored total less duplicate contribution';BoardFeet=$reconciliation.CachedLessDuplicates;Meaning='Diagnostic subtraction; cache aggregation is unverified, so this is not a fresh application total.'}
 )
 $display=@($items | ForEach-Object {
  $basis=switch($_.Calculation){
   'csdproj formula-based estimate' {ArchiveBoardFeetBasis $bdft}
   'DB recalculated on file length basis' {ArchiveBoardFeetBasis $bdft}
   'DB stored project total' {$basisEvidence.CachedBasis}
   'DB stored total less duplicate contribution' {'Unverified'}
   default {$basisEvidence.PieceBasis}
  }
  [pscustomobject]@{Calculation=$_.Calculation;BDFT=(FormatBoardFeet $_.BoardFeet $basis 6);LengthBasis=$basis;Meaning=$_.Meaning}
 })
 $b=[Text.StringBuilder]::new();$null=$b.Append('<details><summary><b>BDFT calculation and duplicate reconciliation</b></summary><p>'+(Enc $bdft.Note)+'</p><p><b>Director length mode:</b> '+(Enc $reconciliation.DatabaseLengthMode)+'</p>'+(Grid $display)+'<p>'+(Enc $basisEvidence.Note)+'</p><p>'+(Enc $reconciliation.Note)+'</p>')
 $null=$b.Append('<p><b>Same-basis difference (file minus recalculated DB):</b> '+$(if($null -eq $fileBasisComparison.Difference){'Unavailable'}else{Enc (([decimal]$fileBasisComparison.Difference).ToString('N6',[Globalization.CultureInfo]::InvariantCulture))})+'</p><p>'+(Enc $fileBasisComparison.Note)+'</p><details><summary>Same-basis comparison by truss</summary>'+(Grid $fileBasisComparison.Trusses)+'</details>')
 $null=$b.Append('<details><summary>Archive truss length settings and quantities</summary>'+(Grid $bdft.Trusses)+'</details>')
 $different=@($reconciliation.Trusses | Where-Object {$_.DuplicateContribution -ne 0 -or [math]::Abs([double]$_.DifferenceAfterDeduplication) -gt 0.01})
 $null=$b.Append('<details><summary>Trusses with duplicate footage or remaining differences above 0.01 BDFT ('+$different.Count+')</summary>'+(Grid $different)+'</details></details>');$b.ToString()
}
function CompareTrussChecksums($zip,$databaseRows){
 $files=@{}
 foreach($entry in $zip.Entries){
  if($entry.FullName -match '(?i)^Trusses/[^/]+\.tdlTruss$'){
   $name=[IO.Path]::GetFileNameWithoutExtension($entry.Name)
   $files[$name]=@($files[$name] | Where-Object {$null -ne $_})+@($entry)
  }
 }
 $dbNames=@{};foreach($record in $databaseRows){$dbNames[$record.Name]=1+$dbNames[$record.Name]}
 foreach($record in $databaseRows){
  $result=[ordered]@{Truss=[string]$record.Name;ComponentKey=[string]$record.ComponentHeaderKey;Entry='';DatabaseChecksum=[string]$record.TrussFileCheckSum;FileMD5='';Result='Unavailable';Note=''}
  $entries=@($files[$record.Name] | Where-Object {$null -ne $_})
  if($dbNames[$record.Name] -gt 1 -or $entries.Count -gt 1){$result.Note='Multiple matching DB records or canonical truss files'}
  elseif($entries.Count -eq 0){$result.Result='Missing file';$result.Note='No canonical Trusses/<name>.tdlTruss entry'}
  else{
   $result.Entry=$entries[0].FullName
   try{
    $md5=[Security.Cryptography.MD5]::Create();$stream=$null
    try{$stream=$entries[0].Open();$result.FileMD5=[BitConverter]::ToString($md5.ComputeHash($stream))}finally{if($stream){$stream.Dispose()};$md5.Dispose()}
    $normalized=$result.DatabaseChecksum.Replace('-','').Trim()
    if($normalized -notmatch '^[0-9a-fA-F]{32}$'){$result.Note='DB checksum missing or not a valid 16-byte MD5'}
    elseif($normalized -ieq $result.FileMD5.Replace('-','')){$result.Result='Same file'}
    else{$result.Result='Different file'}
   }catch{$result.Note=$_.Exception.Message}
  }
  [pscustomobject]$result
 }
}
function FindDuplicatePieces($table){
 # Compare all stored business fields exactly. Identity and audit differences do not
 # make two otherwise identical pieces distinct. ComponentKey scopes each group.
 $excluded=@('TrussName','ComponentKey','PieceKey','CreatedUserKey','LastModifiedUserKey','CreatedDateTime','LastModifiedDateTime')
 $fields=@($table.Columns | ForEach-Object ColumnName | Where-Object {$_ -notin $excluded} | Sort-Object)
 if(-not $fields.Count){throw 'No piece business fields available for duplicate detection'}
 $buckets=[Collections.Generic.Dictionary[string,object]]::new([StringComparer]::Ordinal)
 foreach($piece in $table.Rows){
  $values=[Collections.Generic.List[object]]::new();$values.Add([string]$piece.ComponentKey)
  foreach($field in $fields){
   $v=$piece[$field]
   if($v -is [DBNull]){$values.Add($null)}
   elseif($v -is [double] -or $v -is [single]){$values.Add($v.ToString('R',[Globalization.CultureInfo]::InvariantCulture))}
   elseif($v -is [datetime]){$values.Add($v.Ticks.ToString([Globalization.CultureInfo]::InvariantCulture))}
   elseif($v -is [byte[]]){$values.Add([Convert]::ToBase64String($v))}
   elseif($v -is [IFormattable]){$values.Add($v.ToString($null,[Globalization.CultureInfo]::InvariantCulture))}
   else{$values.Add([string]$v)}
  }
  $signature=ConvertTo-Json -InputObject $values.ToArray() -Compress
  if(-not $buckets.ContainsKey($signature)){$buckets[$signature]=[Collections.Generic.List[object]]::new()}
  $buckets[$signature].Add($piece)
 }
 $groups=[Collections.Generic.List[object]]::new();$affected=@{};$extra=0
 foreach($bucket in $buckets.Values){
  if($bucket.Count -lt 2){continue}
  $first=$bucket[0];$data=[ordered]@{};foreach($field in $fields){$data[$field]=if($first[$field] -is [DBNull]){$null}else{$first[$field]}}
  $records=@(foreach($piece in $bucket){[pscustomobject]@{PieceKey=[string]$piece.PieceKey;CreatedDateTime=RawDate $piece.CreatedDateTime;LastModifiedDateTime=RawDate $piece.LastModifiedDateTime;CreatedUserKey=[string]$piece.CreatedUserKey;LastModifiedUserKey=[string]$piece.LastModifiedUserKey}})
  $groups.Add([pscustomobject]@{Truss=[string]$first.TrussName;ComponentKey=[string]$first.ComponentKey;MatchingRows=$bucket.Count;ExtraPieceRows=$bucket.Count-1;Records=@($records | Sort-Object PieceKey);MatchingData=[pscustomobject]$data})
  $affected[[string]$first.ComponentKey]=$true;$extra+=$bucket.Count-1
 }
 [pscustomobject]@{AffectedTrusses=$affected.Count;ExtraPieceRows=$extra;Groups=@($groups | Sort-Object Truss,ComponentKey,@{Expression={$_.Records[0].PieceKey}});Fields=$fields;Note='Exact matches across all stored piece business fields within each DB truss; suspected duplicates, not confirmed errors.'}
}
function DuplicatePieceDetails($result){
 $b=[Text.StringBuilder]::new()
 $class=if($result.ExtraPieceRows -gt 0){'error'}else{''}
 $null=$b.Append('<details class="'+$class+'"><summary><b>Suspected DB piece duplicates</b></summary><p>'+(Enc $result.Note)+'</p>')
 if($null -eq $result.ExtraPieceRows){$null=$b.Append('<p>Counts unavailable. This does not mean zero duplicates.</p>')}
 elseif($result.ExtraPieceRows -eq 0){$null=$b.Append('<p>No exact duplicate piece rows found under these matching rules.</p>')}
 else{
  $overview=@($result.Groups | ForEach-Object {[pscustomobject]@{Truss=$_.Truss;ComponentKey=$_.ComponentKey;MatchingRows=$_.MatchingRows;ExtraPieceRows=$_.ExtraPieceRows}})
  $null=$b.Append((Grid $overview))
  foreach($group in $result.Groups){
   $null=$b.Append('<details><summary>'+(Enc $group.Truss)+' &mdash; '+$group.MatchingRows+' matching rows, '+$group.ExtraPieceRows+' extra</summary>'+(Grid $group.Records))
   $data=@($group.MatchingData.PSObject.Properties | ForEach-Object {[pscustomobject]@{Field=$_.Name;MatchingValue=$_.Value}})
   $null=$b.Append((Grid $data)+'</details>')
  }
 }
 $null=$b.Append('</details>');$b.ToString()
}
function ReportLegendStyle {
 @'
<style id="report-legend-style">
.report-legend{padding:16px;line-height:1.55}.report-legend>summary{font-weight:650;font-size:16px;color:#243348}.report-legend>summary .legend-summary-hint{font-size:13px;font-weight:400;color:#596a7b;margin-left:10px}.report-legend .legend-start{background:#eef4fa;border-left:4px solid #33465c;padding:12px 16px;margin:16px 0}.report-legend .legend-start p{margin:0 0 8px}.report-legend .legend-start ol{margin:0;padding-left:22px}.report-legend .legend-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px}.report-legend .legend-card{border:1px solid #d9e2eb;border-radius:6px;padding:14px 16px;background:#fff;min-width:0}.report-legend .legend-card h3{font-size:16px;line-height:1.35;margin:0 0 10px;color:#243348}.report-legend .legend-number{display:inline-block;background:#e8eef5;border-radius:4px;padding:1px 7px;margin-right:7px;font-size:13px;color:#33465c}.report-legend .legend-card ul{margin:0;padding-left:20px}.report-legend .legend-card li{margin:6px 0}.report-legend .legend-card p{margin:10px 0 0}.report-legend .legend-key{font-weight:700}.report-legend .legend-card details{padding:9px 11px;margin:12px 0 0;background:#f7f9fb;border-color:#e0e6ed;border-radius:4px}.report-legend .legend-card details summary{font-weight:600;color:#43566b}.report-legend .legend-card details ul{margin-top:8px}.report-legend .legend-tags{display:flex;gap:6px;flex-wrap:wrap;margin:10px 0 0}.report-legend .legend-tag{display:inline-block;padding:3px 8px;border-radius:4px;font-weight:600;font-size:13px}.report-legend .legend-foot{display:flex;flex-wrap:wrap;gap:10px 24px;background:#f3f6f9;border-radius:4px;margin-top:14px;padding:10px 14px}.report-legend .legend-settings{margin:12px 0 0;background:#f7f9fb;border-radius:4px}.report-legend .legend-settings ul{padding-left:20px}.report-legend .legend-settings li{margin:6px 0}.report-legend code{font-size:12px;overflow-wrap:anywhere}.report-legend .legend-fields{columns:2;column-gap:28px}.report-legend .legend-fields li{break-inside:avoid}@media(min-width:1600px){.report-legend .legend-grid{grid-template-columns:repeat(3,minmax(0,1fr))}}@media(max-width:800px){.report-legend .legend-grid{grid-template-columns:1fr}.report-legend>summary .legend-summary-hint{display:block;margin-left:0}.report-legend .legend-fields{columns:1}}
</style>
'@
}
function OverviewStyle {
 @'
<style id="overview-layout-style">
body.overview-page{height:100vh;height:100dvh;box-sizing:border-box;margin:0;padding:12px 16px;overflow:hidden;display:flex;flex-direction:column;gap:10px}
.overview-page>h2{margin:0;font-size:22px;line-height:1.3;flex-shrink:0}
.overview-page>.report-legend{margin:0;padding:12px 16px;flex-shrink:0;min-height:0;max-height:60vh;max-height:min(60dvh,calc(100dvh - 150px));overflow:auto;box-sizing:border-box}
.overview-page>.report-legend>summary{display:flex;align-items:center;gap:8px 12px;flex-wrap:wrap;list-style:none}
.overview-page>.report-legend>summary::-webkit-details-marker{display:none}
.overview-page>.report-legend>summary::before{content:"";border-top:5px solid transparent;border-bottom:5px solid transparent;border-left:7px solid #243348;flex-shrink:0}
.overview-page>.report-legend[open]>summary::before{transform:rotate(90deg)}
.overview-page .legend-summary-label{min-width:0}
.overview-page .legend-run-context{margin-left:auto;text-align:right;font-size:13px;font-weight:400;color:#536475;overflow-wrap:anywhere}
.overview-page>.overview-table{flex:1;min-height:0;max-height:none;margin:0;overflow:auto}
@media(max-width:800px){body.overview-page{padding:10px}.overview-page>h2{font-size:20px}.overview-page .legend-run-context{flex-basis:100%}}
</style>
'@
}
function ReportLegend($fileZoneName,$databaseZoneName,[double]$timestampTolerance,[double]$numberTolerance,[bool]$includeTimestamp=[bool]$Timestamp,[string]$context=''){
 $fileLabel=Enc $fileZoneName;$databaseLabel=Enc $databaseZoneName
 $secondsLabel=Enc ($timestampTolerance.ToString('G',[Globalization.CultureInfo]::InvariantCulture))
 $numberLabel=Enc ($numberTolerance.ToString('G',[Globalization.CultureInfo]::InvariantCulture))
 $contextHtml=if($context){'<span class="legend-run-context">'+(Enc $context)+'</span>'}else{''}
 $html=@"
<details class="report-legend"><summary><span class="legend-summary-label">How to read this report <span class="legend-summary-hint">Column guide, colors and comparison rules</span></span>$contextHtml</summary>
<div class="legend-start"><p><strong>Start with the three black column headers: In both, Field sync comparison / Same, and Truss file checksum / Same.</strong></p><ol><li><strong>Check truss inventory.</strong> The tdlTruss total, DB and In both should be equal if both sources contain the same truss names.</li><li><strong>Check the matched data.</strong> Same should equal In both. If In both is zero, there are no matched trusses to compare.</li><li><strong>Check file contents, board footage and duplicates.</strong> These are separate checks. Use timestamps to understand the sequence of events.</li></ol><p style="margin:10px 0 0"><strong>Status:</strong> inventory, checked fields and recorded checksums are reported separately from BDFT. Healthy additionally requires comparable footage agreement per truss and in total, matching quantities, and no suspected duplicate piece rows. A different stored Director total or length method alone does not prevent Healthy status; a yellow BDFT notice remains visible for that difference. Equal totals do not prove that every piece is identical.</p><p style="margin:8px 0 0"><strong>A green cell confirms only that check.</strong> It does not prove that the whole project is in sync. A difference means review is needed; it does not establish its cause.</p></div>
<div class="legend-grid">
<section class="legend-card"><h3><span class="legend-number">1</span>Truss inventory &mdash; what is present?</h3><ul><li><span class="legend-key">csdproj / JSON:</span> the number of truss JSON files stored in the csdproj. These can include files retained after a truss is removed.</li><li><span class="legend-key">csdproj / tdlTruss:</span> the number of main truss files stored in the csdproj. This count is used for matching to the DB.</li><li><span class="legend-key">DB:</span> the number of truss/component records found for that project.</li><li><span class="legend-key">In both:</span> truss names found in both the tdlTruss files and the DB. Matching uses tdlTruss, not JSON.</li></ul><div class="legend-tags"><span class="legend-tag close">Green: tdlTruss, DB and In both equal</span><span class="legend-tag older">Amber: JSON and tdlTruss counts differ</span><span class="legend-tag error">Red: counts differ</span><span class="legend-tag unknown">Gray: cannot confirm</span></div><p>JSON files can remain after trusses are removed. Their count is informational and can be higher than the tdlTruss count.</p><details><summary>How names are matched</summary><ul><li>Matches are confined to the project identified by its folder name.</li><li>Only the main <code>Trusses/&lt;name&gt;.tdlTruss</code> files count as truss files. Copies in backup or temporary folders are excluded.</li><li>Names are compared without regard to letter case. Spaces and punctuation are preserved.</li><li>Duplicate names or incomplete processing prevent a confident comparison. Layout/settings JSON is excluded from the JSON truss count.</li><li>Missing from csdproj means a DB truss has no corresponding main tdlTruss file, even if an old JSON file exists.</li></ul></details></section>
<section class="legend-card"><h3><span class="legend-number">2</span>Field sync comparison &mdash; do checked fields agree?</h3><ul><li><span class="legend-key">Same:</span> matched trusses with no differences in the available fields checked.</li><li><span class="legend-key">Different:</span> matched trusses with at least one checked field that differs.</li><li>This group covers trusses <b>in both sources only</b>. Trusses missing from either source are not counted here.</li></ul><div class="legend-tags"><span class="legend-tag close">Green: every matched truss is in Same</span></div><p>This checks selected properties, not every part of a truss. Missing field values are skipped; unresolved comparisons prevent the green highlight.</p><details><summary>Which 19 fields are checked?</summary><ul class="legend-fields"><li>Quantity and number of plies</li><li>Thickness and overall truss height</li><li>Left and right overhang</li><li>Left and right cantilever</li><li>Left and right heel height</li><li>Left and right top-chord pitch</li><li>Left and right bottom-chord pitch</li><li>Attic, gable, girder and flipped flags</li><li>On-center spacing (OCSpacing)</li></ul><p>Timestamps and review markers are checked separately. This group does not compare span, detailed geometry, joints, lumber or plates.</p><p>Numeric differences up to <code>$numberLabel</code> are treated as equal. Missing/null values are skipped and recorded. If no fields can be compared, the truss is unresolved.</p></details></section>
<section class="legend-card"><h3><span class="legend-number">3</span>Truss file checksum &mdash; do file contents agree?</h3><p style="margin:0 0 8px">A checksum is a fingerprint of a file's contents. This compares each main tdlTruss file with the fingerprint recorded in the DB.</p><ul><li><span class="legend-key">Same / Different:</span> the fingerprints agree / differ.</li><li><span class="legend-key">File Missing:</span> a truss exists in the DB but its main tdlTruss file is missing from the csdproj.</li><li><span class="legend-key">Not checked:</span> a missing/invalid DB checksum, duplicate truss names or a file read error prevented comparison.</li></ul><div class="legend-tags"><span class="legend-tag close">Green: Same equals nonzero In both</span><span class="legend-tag error">Red: different or missing files</span><span class="legend-tag unknown">Gray: not checked</span></div><p>Partial matches are not highlighted green. Matching file contents do not establish that related DB piece records are correct.</p><details><summary>Checksum comparison rules</summary><ul><li>MD5 is calculated from the uncompressed bytes of each main <code>.tdlTruss</code> entry and compared with <code>ComponentHeader.TrussFileCheckSum</code>.</li><li>The ZIP and JSON files are not the files hashed for this check.</li><li>All DB component records are assessed, including those without a matching JSON file.</li><li>Missing/invalid DB checksums, duplicate names or read errors leave a comparison Not checked. Unresolved comparisons prevent the Same green highlight.</li></ul></details></section>
<section class="legend-card"><h3><span class="legend-number">4</span>BDFT &mdash; compare the same length basis</h3><ul><li><span class="legend-key">csdproj:</span> archive lumber, piece lengths and plies, using each truss's saved length setting and Studio.json layout quantities.</li><li><span class="legend-key">DB (calculated):</span> freshly calculated DB piece geometry using the same length modes and layout quantities. This is the directly comparable value, not a stored Director total.</li><li><span class="legend-key">Difference:</span> file formula minus DB formula. Nonzero means the geometry-based footage differs even after aligning length modes and quantities. Individual truss differences are checked too, so offsetting errors cannot establish agreement.</li><li><span class="legend-key">DB (As stored):</span> Director's unchanged cached project total, shown for reference with its length-method marker where verified. It is not used in the Difference column and may reflect another method or older cached data.</li></ul><p><b>* = Pick Length.</b> Unmarked source totals use Actual Length, or both stored length calculations give identical footage. <b>(Mixed)</b> means both length modes contribute. <b>(basis unverified)</b> means there is insufficient evidence to identify the total's length method: required data may be missing or the stored total may not reconcile with its piece records. An unverified basis does not by itself mean the footage is wrong. An em dash means the total is unavailable. Difference = csdproj minus DB (calculated). It has no length marker.</p><details><summary>Calculation rules and missing values</summary><ul><li>One-ply BDFT = thickness &times; width &times; length in inches &divide; 144. Multiply by piece plies, then layout quantity. SQL decimal rounding is reproduced. No extra multiplication by truss plies is applied.</li><li>Dimensional lumber uses nominal size (e.g. 2&times;4); engineered lumber uses actual catalog dimensions. Actual Length refers to piece length, not a switch to actual dimensional-lumber thickness/width.</li><li>The file's <code>OtheruseExactBdft=true</code> selects Overall (Actual Length); false selects PickLength. Only when the truss setting is absent does the archive environment's <code>useExactBdft</code> select the mode. Missing/invalid settings make the total unavailable.</li><li>Missing/invalid required lengths make the calculation unavailable. No alternate length is substituted. A saved zero length is distinct from a missing value and contributes zero.</li><li>Both formula totals use geometry only: Director per-unit pricing exceptions are excluded on both sides. These are not predictions of an application cache refresh. The DB stored total is preserved for that separate check.</li><li>Both formulas use file layout membership and quantities. DB-only trusses remain in the inventory check; DB quantity differences are reported separately. Retained JSON outside the layouts is excluded. Duplicate DB pieces are not removed.</li><li>DB stored-total markers require reconciliation with its piece caches, piece plies and DB quantities (0.000001 BDFT tolerance); the current preset alone cannot label an older cache. This verifies agreement with stored length-specific caches, not their freshness or historical provenance.</li><li>Project details show each truss's selected mode, setting source, quantities and calculated footage. Partial totals are withheld when required evidence is missing.</li></ul></details></section>
<section class="legend-card"><h3><span class="legend-number">5</span>DB row duplicates &mdash; is piece data repeated?</h3><ul><li><span class="legend-key">Affected trusses:</span> DB trusses containing repeated piece data.</li><li><span class="legend-key">Extra piece rows:</span> repetitions beyond the first copy. Three identical rows count as two extra rows.</li><li>Counts cover all DB trusses in the project, including trusses missing from the csdproj.</li></ul><div class="legend-tags"><span class="legend-tag error">Red: suspected duplicates found</span></div><p>These counts refer to suspected duplicate piece rows within a DB truss, not duplicates across every DB table. They are candidates for review, not confirmed errors. The project page shows the matching records and their identifiers.</p><details><summary>What makes two piece rows a suspected duplicate?</summary><ul><li>They belong to the same DB truss (ComponentKey), and all stored business values match exactly: geometry, engineering label, lumber/pricing references, dimensions, lengths, plies and calculated footage/cost values.</li><li>Record IDs and creation/modification dates and users are excluded from matching.</li><li>Comparisons are case-sensitive and use no numeric tolerance. Pieces with differing business values are not counted.</li><li>Counts identify distinct truss component IDs and repeat for each archive of the same DB project. Legitimate repetitions require manual review.</li><li>Missing/ambiguous DB projects or a failed piece query produce unavailable counts, not zero duplicates.</li></ul></details></section>
<section class="legend-card"><h3><span class="legend-number">6</span>Timestamps &mdash; which date is earlier?</h3><ul><li><span class="legend-key">Older:</span> the JSON timestamp inside the csdproj is earlier than the DB truss timestamp by more than $secondsLabel seconds.</li><li><span class="legend-key">Close:</span> the two timestamps are within $secondsLabel seconds.</li><li><span class="legend-key">Newer:</span> the JSON timestamp is later than the DB truss timestamp by more than $secondsLabel seconds.</li><li><span class="legend-key">Minimum / maximum difference:</span> the smallest / largest signed JSON-minus-DB time gap. <b>&minus;2d</b> means two days earlier; <b>+2d</b> means two days later.</li></ul><div class="legend-tags"><span class="legend-tag older">Amber: older</span><span class="legend-tag close">Green: close</span><span class="legend-tag newer">Blue: newer</span><span class="legend-tag mixed">Purple: both older and newer</span><span class="legend-tag unknown">Gray: unavailable</span></div><p>Dates are investigation clues. They do not prove an older/newer truss version or explain the cause of a difference.</p><details><summary>Timestamp comparison rules</summary><ul><li>These counts cover matched trusses only and compare ZIP-entry JSON timestamps with <code>ComponentTruss.LastModifiedDateTime</code>, not the outer csdproj file's save date.</li><li>Timestamp counts overlap field-comparison counts. Purple in Status means some matched JSON timestamps are older and others are newer.</li><li>Missing JSON, missing DB times, or ambiguous/invalid daylight-saving times make the comparison unavailable.</li><li>Signed difference colors indicate earlier/later even within tolerance. Processing errors take priority over timestamp hints.</li><li>Packaging can affect how timestamps are represented. Review the original archive metadata before using dates as root-cause evidence.</li></ul></details></section>
<section class="legend-card"><h3><span class="legend-number">7</span>Review markers &mdash; a separate review check</h3><ul><li><span class="legend-key">Review markers differ:</span> matched trusses whose CSEngineer review/sealing marker differs between JSON and the DB.</li><li>This marker is assigned by the external review program. It is not the file checksum and does not by itself establish a design difference.</li></ul><details><summary>Technical field</summary><p>The value compared is <code>TrussMatchCode</code>. Marker-only differences are counted in the summary but excluded from the main difference detail list.</p></details></section>
<section class="legend-card"><h3><span class="legend-number">8</span>Status and unresolved comparisons &mdash; what needs review?</h3><ul><li><span class="legend-key">Status:</span> reports inventory/field/file agreement alongside a separate BDFT result, and highlights missing trusses, processing problems and timestamp direction. Several messages may appear together.</li><li><span class="legend-key">Unresolved comparisons:</span> trusses that could not be compared reliably. These are not confirmed data differences.</li><li><span class="legend-key">Unavailable / Unknown:</span> insufficient information to complete that check. This does not mean the values agree.</li></ul><details><summary>Common reasons and overview sections</summary><ul><li>Duplicate JSON/tdlTruss names, multiple matching DB trusses or project records, missing required JSON, or no usable checked fields can leave a comparison unresolved.</li><li>Parsing and processing errors are reported separately. Review the project page for the reason.</li><li><b>Results:</b> projects with comparison data or issues.</li><li><b>Project not in DB:</b> no DB project matches the folder's identifier, even if the file counts are zero.</li><li><b>Empty:</b> projects with no trusses to compare.</li></ul></details></section>
</div>
<p><b>Report rows:</b> each row represents one project and one csdproj archive. Multiple archives for a project stay separate to preserve their saved states. Click a project name to open its details.</p><div class="legend-foot"><span><b>0</b> = an actual count of zero</span><span><b>&mdash;</b> = unavailable or no comparison could be made</span><span>Open a project to review the supporting details.</span></div>
<details class="legend-settings"><summary>Technical settings for this run</summary><ul><li>Timestamp tolerance: <b>$secondsLabel seconds</b>.</li><li>Numeric field tolerance: <code>$numberLabel</code>.</li><li>csdproj JSON timestamp timezone: <b>$fileLabel</b>.</li><li>DB timestamp timezone: <b>$databaseLabel</b>.</li><li>Both timezone interpretations default to the machine's current timezone when the script runs. <code>-FileTimeZone</code> and <code>-DatabaseTimeZone</code> can override them independently.</li><li>Processing is read-only. Original csdproj files and DB records are unchanged; DeletedProjects folders and Attachment_/Attachments_ archives are skipped.</li></ul></details>
</details>
"@
 if(-not $includeTimestamp){
  $html=[regex]::Replace($html,'(?s)<section class="legend-card"><h3><span class="legend-number">6</span>Timestamps.*?</section>','')
  $html=$html.Replace('Use timestamps to understand the sequence of events.','Timestamp comparison is optional; enable it with <code>-Timestamp</code>.')
  $html=$html.Replace('Timestamps and review markers are checked separately.','Review markers are checked separately. Timestamp comparison is optional.')
  $html=$html.Replace('highlights missing trusses, processing problems and timestamp direction.','highlights missing trusses and processing problems.')
  foreach($pattern in @('<li>Timestamp tolerance:.*?</li>','<li>csdproj JSON timestamp timezone:.*?</li>','<li>DB timestamp timezone:.*?</li>','<li>Both timezone interpretations.*?</li>')){$html=[regex]::Replace($html,$pattern,'')}
  $html=$html.Replace('<span class="legend-number">7</span>','<span class="legend-number">6</span>').Replace('<span class="legend-number">8</span>','<span class="legend-number">7</span>')
  $html=$html.Replace('<summary>Technical settings for this run</summary><ul>','<summary>Technical settings for this run</summary><ul><li>Timestamp comparison: <b>off</b>. Enable with <code>-Timestamp</code>.</li>')
 }
 $html
}
function SummaryHeader {
 $header='<table class="grouped-summary"><thead><tr><th rowspan="3">Project</th><th rowspan="3">csdproj</th><th rowspan="3" class="status-column">Status</th><th colspan="4" style="background:#e3edf7">Truss inventory</th><th colspan="2" style="background:#e8eef0">Field sync comparison</th><th colspan="4" style="background:#e4ecf0">Truss file checksum</th><th colspan="4" style="background:#e6eee1">BDFT comparison</th><th colspan="2" style="background:#f4e6e4">DB row duplicates</th><th colspan="6" style="background:#eee8f6">Timestamps - in both only</th><th rowspan="3">Review markers differ</th><th rowspan="3">Unresolved comparisons</th></tr><tr><th colspan="2" scope="colgroup">csdproj</th><th rowspan="2" scope="col">DB</th><th rowspan="2" scope="col" class="key-inventory">In both</th><th rowspan="2" class="key-data">Same</th><th rowspan="2">Different</th><th rowspan="2" class="key-checksum">Same</th><th rowspan="2">Different</th><th rowspan="2">File Missing</th><th rowspan="2">Not checked</th><th rowspan="2">csdproj</th><th rowspan="2">DB (calculated)</th><th rowspan="2">Difference</th><th rowspan="2">DB (As stored)</th><th rowspan="2">Affected trusses</th><th rowspan="2">Extra piece rows</th><th rowspan="2">Older</th><th rowspan="2">Close</th><th rowspan="2">Newer</th><th rowspan="2">Unavailable</th><th rowspan="2">Minimum difference</th><th rowspan="2">Maximum difference</th></tr><tr><th scope="col">JSON</th><th scope="col">tdlTruss</th></tr></thead><tbody>'
 if(-not $Timestamp){
  $header=$header.Replace('<th colspan="6" style="background:#eee8f6">Timestamps - in both only</th>','').Replace('<th rowspan="2">Older</th><th rowspan="2">Close</th><th rowspan="2">Newer</th><th rowspan="2">Unavailable</th><th rowspan="2">Minimum difference</th><th rowspan="2">Maximum difference</th>','')
 }
 $header
}
function SummaryRow($row,$href='') {
 $paired=[int]$row.TdlMatchedNames
 $uncertain=$row.Ambiguous -gt 0 -or $row.Status -match 'errors|incomplete'
 $inventory=if($null -eq $row.TdlTrussCount){@('Unknown','Unknown','Unknown','Unknown')}else{@([int]$row.JsonTrussCount,[int]$row.TdlTrussCount,[int]$row.DbTrussCount,$paired)}
 if($row.Status -eq 'Project not in DB'){$inventory[2]='-';$inventory[3]='-'}
 $project=Enc $row.Project
 if($href){$project='<a href="'+(Enc $href)+'">'+$project+'</a>'}
 $status=StatusHtml $row
 if(-not $uncertain -and $paired -eq 0 -and $row.Status -ne 'Project not in DB'){
  if($row.'Only in csdproj' -gt 0 -or $row.'Only in database' -gt 0){$status='<div class="error" style="padding:5px">No matching truss names - data comparison not possible</div>'+$status}
  else{$status='No trusses in either source'}
 }
 $b=[Text.StringBuilder]::new();$null=$b.Append('<tr><td>'+$project+'</td><td>'+(Enc $row.csdproj)+'</td><td class="status-column">'+$status+'</td>')
 for($i=0;$i -lt $inventory.Count;$i++){
  $class=''
  if($i -eq 0 -and $null -ne $row.JsonTrussCount -and $null -ne $row.TdlTrussCount -and $row.JsonTrussCount -ne $row.TdlTrussCount){$class='older'}
  if($i -eq 3){$class=if($uncertain){'unknown'}elseif($row.Status -eq 'Project not in DB'){'unknown'}elseif($row.TdlTrussCount -eq $row.DbTrussCount -and $row.DbTrussCount -eq $paired){'close'}else{'error'}}
  $null=$b.Append('<td class="'+$class+'">'+(Enc $inventory[$i])+'</td>')
 }
 foreach($col in @('Data synced','Data differs','ChecksumSame','ChecksumDifferent','ChecksumMissing','ChecksumUnavailable','CsdprojBoardFeet','DatabaseOnFileBasisBoardFeet','SameBasisBoardFeetDifference','DatabaseBoardFeet','DuplicateAffectedTrusses','DuplicateExtraPieceRows','Older Timestamp','Within tolerance','Newer Timestamp','Unavailable','Minimum difference','Maximum difference','Review marker differs','Ambiguous')){
  if(-not $Timestamp -and $col -in @('Older Timestamp','Within tolerance','Newer Timestamp','Unavailable','Minimum difference','Maximum difference')){continue}
  $value=Enc $row.$col;$class=CellColor $col $row.$col
  if($col -eq 'ChecksumSame'){$class=if(-not $uncertain -and $paired -gt 0 -and [int]$row.ChecksumSame -eq $paired){'close'}else{''}}
  if($col -in @('ChecksumDifferent','ChecksumMissing','ChecksumUnavailable') -and $row.$col -gt 0){$class=switch($col){'ChecksumDifferent'{'error'} 'ChecksumMissing'{'error'} default{'unknown'}}}
  if($col -eq 'Data synced' -and -not $uncertain -and $paired -gt 0 -and [int]$row.'Data synced' -eq $paired){$class='close'}
  if($col -in @('CsdprojBoardFeet','DatabaseOnFileBasisBoardFeet','SameBasisBoardFeetDifference','DatabaseBoardFeet','BoardFeetDifference')){
   if($col -in @('BoardFeetDifference','SameBasisBoardFeetDifference')){$value=if($null -eq $row.$col){'&mdash;'}else{Enc (([double]$row.$col).ToString('N2',[Globalization.CultureInfo]::InvariantCulture))}}
   else{$basis=if($col -in @('CsdprojBoardFeet','DatabaseOnFileBasisBoardFeet')){$row.CsdprojBoardFeetBasis}else{$row.DatabaseBoardFeetBasis};$value=Enc (FormatBoardFeet $row.$col $basis)}
   $class=if($null -eq $row.$col){'unknown'}else{''}
  }
  if($col -in @('DuplicateAffectedTrusses','DuplicateExtraPieceRows')){$value=if($null -eq $row.$col){'&mdash;'}else{Enc $row.$col};$class=if($null -eq $row.$col){'unknown'}elseif($row.$col -gt 0){'error'}else{''}}
  if($paired -eq 0 -and $col -notin @('Only in database','Only in csdproj','Ambiguous','ChecksumSame','ChecksumDifferent','ChecksumMissing','ChecksumUnavailable','CsdprojBoardFeet','DatabaseOnFileBasisBoardFeet','SameBasisBoardFeetDifference','DatabaseBoardFeet','DuplicateAffectedTrusses','DuplicateExtraPieceRows')){$value='&mdash;';$class='unknown'}
  $null=$b.Append('<td class="'+$class+'">'+$value+'</td>')
 }
 $null=$b.Append('</tr>');$b.ToString()
}
function SummaryTable($rows){$b=[Text.StringBuilder]::new();$null=$b.Append((SummaryHeader));foreach($row in $rows){$null=$b.Append((SummaryRow $row))};$null=$b.Append('</tbody></table>');$b.ToString()}

function ConsoleStatsColumns([string[]]$projectNames,[bool]$includeTimestamp){
 $projectWidth=10
 foreach($name in $projectNames){$projectWidth=[math]::Max($projectWidth,[math]::Min(20,$name.Length))}
 $columns=@(
  [pscustomobject]@{Name='Project';Width=$projectWidth;Right=$false}
  [pscustomobject]@{Name='Archive';Width=7;Right=$false}
  [pscustomobject]@{Name='JSON';Width=5;Right=$true}
  [pscustomobject]@{Name='tdlTruss';Width=8;Right=$true}
  [pscustomobject]@{Name='DB';Width=5;Right=$true}
  [pscustomobject]@{Name='In both';Width=7;Right=$true}
  [pscustomobject]@{Name='Fields S/D';Width=11;Right=$true}
  [pscustomobject]@{Name='Checksum S/D/M/N';Width=16;Right=$true}
  [pscustomobject]@{Name='ExtraRows';Width=9;Right=$true}
  [pscustomobject]@{Name='BDFT same diff';Width=14;Right=$true}
 )
 if($includeTimestamp){$columns+= [pscustomobject]@{Name='Time O/C/N/U';Width=13;Right=$true}}
 $columns+= [pscustomobject]@{Name='Status';Width=24;Right=$false}
 $columns
}
function ConsoleStatsLines($columns,[object[]]$values){
 # Wrap long names/statuses within their column instead of dropping information.
 $chunks=@();$height=1
 for($i=0;$i -lt $columns.Count;$i++){
  $text=([string]$values[$i]) -replace '[\x00-\x1f\x7f]',' '
  $parts=[Collections.Generic.List[string]]::new()
  if(-not $text){$parts.Add('')}
  for($offset=0;$offset -lt $text.Length;$offset+=$columns[$i].Width){$parts.Add($text.Substring($offset,[math]::Min($columns[$i].Width,$text.Length-$offset)))}
  $chunks+=,@($parts);$height=[math]::Max($height,$parts.Count)
 }
 for($line=0;$line -lt $height;$line++){
  $cells=@(for($i=0;$i -lt $columns.Count;$i++){
   $part=if($line -lt $chunks[$i].Count){$chunks[$i][$line]}else{''}
   if($columns[$i].Right){$part.PadLeft($columns[$i].Width)}else{$part.PadRight($columns[$i].Width)}
  })
  $cells -join ' '
 }
}
function ConsoleStatsValues($row,[int]$archiveNumber,[int]$archiveTotal,[bool]$includeTimestamp){
 $knownInventory=$null -ne $row.TdlTrussCount
 $knownDb=$row.Status -ne 'Project not in DB'
 $paired=[int]$row.TdlMatchedNames
 $status=[string]$row.Status
 if($knownInventory -and $knownDb -and $row.TdlTrussCount -eq 0 -and $row.DbTrussCount -eq 0 -and $status -eq 'Compared'){$status='Empty'}
 elseif($status -eq 'Compared'){$status='Needs review'}
 $values=@(
  $row.Project
  ($archiveNumber.ToString()+'/'+$archiveTotal)
  $(if($knownInventory){$row.JsonTrussCount}else{'-'})
  $(if($knownInventory){$row.TdlTrussCount}else{'-'})
  $(if($knownDb){$row.DbTrussCount}else{'-'})
  $(if($knownInventory -and $knownDb){$paired}else{'-'})
  $(if($paired -gt 0){[string]$row.'Data synced'+'/'+$row.'Data differs'}else{'-'})
  $(if($knownDb){[string]$row.ChecksumSame+'/'+$row.ChecksumDifferent+'/'+$row.ChecksumMissing+'/'+$row.ChecksumUnavailable}else{'-'})
  $(if($null -ne $row.DuplicateExtraPieceRows){$row.DuplicateExtraPieceRows}else{'-'})
  $(if($null -ne $row.SameBasisBoardFeetDifference){([decimal]$row.SameBasisBoardFeetDifference).ToString('+0.00;-0.00;0.00',[Globalization.CultureInfo]::InvariantCulture)}else{'-'})
 )
 if($includeTimestamp){$values+= $(if($paired -gt 0){[string]$row.'Older Timestamp'+'/'+$row.'Within tolerance'+'/'+$row.'Newer Timestamp'+'/'+$row.Unavailable}else{'-'})}
 $values+= $status
 $values
}
function WriteConsoleStatsHeader($columns,[bool]$includeTimestamp){
 Write-Host ''
 Write-Host 'One row per csdproj archive; Archive is the saved archive number within its project.'
 Write-Host 'Fields: Same/Different (matched trusses). Checksum: Same/Different/File Missing/Not checked (DB trusses).'
 Write-Host 'ExtraRows: suspected duplicate DB piece rows. BDFT same diff: file formula minus DB formula using file modes and layout quantities. - means unavailable or no comparison.'
 if($includeTimestamp){Write-Host 'Time: Older/Close/Newer/Unavailable JSON timestamps (matched trusses).'}
 Write-Host ''
 foreach($line in (ConsoleStatsLines $columns @($columns | ForEach-Object {$_.Name}))){Write-Host $line -ForegroundColor Cyan}
 Write-Host ((@($columns | ForEach-Object {'-' * $_.Width})) -join ' ') -ForegroundColor DarkGray
}
function WriteConsoleStatsRow($columns,$row,[int]$archiveNumber,[int]$archiveTotal,[bool]$includeTimestamp){
 $values=@(ConsoleStatsValues $row $archiveNumber $archiveTotal $includeTimestamp)
 foreach($line in (ConsoleStatsLines $columns $values)){Write-Host $line}
}

function Grid($rows){
 $items=@($rows);if(-not $items.Count){return ''}
 $cols=@($items[0].PSObject.Properties.Name)
 $b=[Text.StringBuilder]::new();$null=$b.Append('<table><tr>')
 foreach($col in $cols){$null=$b.Append('<th>'+(Enc $col)+'</th>')};$null=$b.Append('</tr>')
 foreach($item in $items){$null=$b.Append('<tr>');foreach($col in $cols){$class=CellColor $col $item.$col;$value=Enc $item.$col;if($col -eq 'Status'){$class='';$value=StatusHtml $item};$null=$b.Append('<td class="'+$class+'">'+$value+'</td>')};$null=$b.Append('</tr>')}
 $null=$b.Append('</table>');$b.ToString()
}
$root=(Get-Item -LiteralPath $ProjectsRoot).FullName.TrimEnd('\')
$fileZone=$null;if($Timestamp){$fileZone=[TimeZoneInfo]::FindSystemTimeZoneById($FileTimeZone)}
$dbZone=$null;if($Timestamp){$dbZone=[TimeZoneInfo]::FindSystemTimeZoneById($DatabaseTimeZone)}
$out=[IO.Path]::GetFullPath($OutputDirectory)
$runId=(Get-Date).ToUniversalTime().ToString('yyyyMMddTHHmmssfffZ')+'-'+[guid]::NewGuid().ToString('N').Substring(0,8)
$out=$out.TrimEnd('\','/')+'-'+$runId
$run=$out
$pages=Join-Path $run 'projects'
$null=New-Item -ItemType Directory -Path $pages -Force
$scratchRoot=$null
if(-not $retainEvidence){
 $scratchRoot=Join-Path ([IO.Path]::GetTempPath()) ('CsdprojComparison-'+[guid]::NewGuid().ToString('N'))
 $null=New-Item -ItemType Directory -Path $scratchRoot
}
$style='<style>body{font:14px Segoe UI,Arial;color:#243348;background:#f5f7fa;margin:28px}h1{margin-bottom:8px}p{line-height:1.5}table{border-collapse:collapse;background:white;width:100%;margin:16px 0}td,th{border:1px solid #dce3eb;padding:9px;text-align:left;vertical-align:top}th{background:#e8eef5}a{color:#175d9b}details{background:white;padding:12px;margin:9px 0;border:1px solid #dde3ea}summary{cursor:pointer}small{color:#536475}.notice{background:#fff2d6;padding:14px}.scroll{overflow:auto}code{overflow-wrap:anywhere}.older{background:#fff0d5;color:#814900}.newer{background:#e0efff;color:#174e83}.mixed{background:#eee4ff;color:#623c87}.close{background:#e0f2e7;color:#245d3b}.unknown{background:#edf0f3;color:#55616e}.error{background:#ffe1df;color:#982a24}/* Sticky headers within scrollable report tables. */
.scroll{max-height:75vh;overflow:auto;position:relative}
.scroll table{margin:0;border-collapse:separate;border-spacing:0}
.scroll th{position:sticky;top:0;z-index:3;background:#e8eef5;box-shadow:0 1px 0 #c8d2df}
.grouped-summary thead tr:first-child th{height:44px;box-sizing:border-box}
.grouped-summary thead tr:nth-child(2) th{top:44px;z-index:2}.grouped-summary thead tr:nth-child(3) th{top:88px;z-index:2}.grouped-summary .status-column{min-width:200px;box-sizing:border-box}
.grouped-summary thead th[rowspan]{z-index:4;vertical-align:middle}body{margin:16px}.overview-table{max-height:82vh}.section-divider td{background:#33465c;color:white;font-weight:700;text-align:left;padding:10px 12px;border-color:#33465c}.scroll th.key-inventory,th.key-inventory,.scroll th.key-data,th.key-data,.scroll th.key-checksum,th.key-checksum{background:#111111;color:#ffffff;font-weight:700;box-shadow:inset 0 -3px 0 #111111}</style>'
$style+=(ReportLegendStyle)
$legend=ReportLegend $FileTimeZone $DatabaseTimeZone $ToleranceSeconds $NumericTolerance
$stickyHeaderScript=@'
<script>(function(){function sizeHeaders(){document.querySelectorAll(".grouped-summary").forEach(function(t){var rows=Array.from(t.querySelectorAll("thead tr")),offset=0;rows.forEach(function(row,index){Array.from(row.children).forEach(function(cell){cell.style.top=offset+"px";cell.style.zIndex=String(rows.length-index+3)});offset+=row.getBoundingClientRect().height})})}window.addEventListener("resize",sizeHeaders);window.addEventListener("load",sizeHeaders);sizeHeaders()})();</script>
'@
function Page($title,$body,$path,[switch]$Overview){
 $heading=if($Overview){'h2'}else{'h1'}
 $bodyAttribute=if($Overview){' class="overview-page"'}else{''}
 $layoutStyle=if($Overview){OverviewStyle}else{''}
 ('<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>'+(Enc $title)+'</title>'+$style+$layoutStyle+'</head><body'+$bodyAttribute+'><'+$heading+'>'+(Enc $title)+'</'+$heading+'>'+$body+$stickyHeaderScript+'</body></html>') | Set-Content -LiteralPath $path -Encoding UTF8
}
# Traverse explicitly so excluded directories are never entered.
$pending=[Collections.Generic.Stack[string]]::new()
$pending.Push($root)
$found=[Collections.Generic.List[object]]::new()
while($pending.Count -gt 0){
 $directory=$pending.Pop()
 foreach($item in Get-ChildItem -LiteralPath $directory){
  if($item.PSIsContainer){
   if($item.Name -ine 'DeletedProjects' -and -not ($item.Attributes -band [IO.FileAttributes]::ReparsePoint)){$pending.Push($item.FullName)}
  }elseif($item.Extension -ieq '.csdproj' -and $item.Name -notmatch '(?i)^Attachments?_'){$found.Add($item)}
 }
}
$archives=@($found | Sort-Object FullName)
$groups=@{}
foreach($a in $archives){$rel=$a.FullName.Substring($root.Length).TrimStart('\');$parts=$rel -split '[\\/]';$id=if($parts.Count -gt 1){$parts[0]}else{'[No project folder]'};if($ProjectId -and $id -notin $ProjectId){continue};$groups[$id]=@($groups[$id] | Where-Object {$null -ne $_})+@($a)}
$selectedIds=@($groups.Keys | Sort-Object)
if($Limit -gt 0){$selectedIds=@($selectedIds | Select-Object -First $Limit)}
$conn=[Data.SqlClient.SqlConnection]::new()
$builder=[Data.SqlClient.SqlConnectionStringBuilder]::new();$builder['Data Source']=$Server;$builder['Initial Catalog']=$Database;$builder['Integrated Security']=$true;$builder['Connect Timeout']=15;$conn.ConnectionString=$builder.ConnectionString
$allSummary=[Collections.Generic.List[object]]::new();$links=[Collections.Generic.List[object]]::new()
try{
 $conn.Open()
 $consoleStatsColumns=@(ConsoleStatsColumns $selectedIds ([bool]$Timestamp))
 WriteConsoleStatsHeader $consoleStatsColumns ([bool]$Timestamp)
 $projectIndex=0
 foreach($id in $selectedIds){
  $projectIndex++;Write-Progress -Id 1 -Activity 'Comparing projects' -Status "[$projectIndex/$($selectedIds.Count)] $id" -PercentComplete (($projectIndex-1)*100/[math]::Max(1,$selectedIds.Count))
  $cmd=$conn.CreateCommand();$cmd.CommandTimeout=60
  $cmd.CommandText='SELECT EventKey,CachedBoardFeet FROM dbo.Project WHERE ProjectNumber=@p';$null=$cmd.Parameters.Add('@p',[Data.SqlDbType]::NVarChar,256);$cmd.Parameters['@p'].Value=$id
  $pt=[Data.DataTable]::new();$reader=$cmd.ExecuteReader();$pt.Load($reader);$reader.Dispose()
  $cmd.CommandText='SELECT h.Name,h.ComponentHeaderKey,h.TrussFileCheckSum,h.LastModifiedDateTime AS HeaderModified,c.LastModifiedDateTime AS ComponentModified,c.ComponentQuantity,t.* FROM dbo.Project p JOIN dbo.ComponentHeader h ON h.ProjectEventKey=p.EventKey LEFT JOIN dbo.Component c ON c.ComponentKey=h.ComponentHeaderKey LEFT JOIN dbo.ComponentTruss t ON t.ComponentKey=c.ComponentKey WHERE p.ProjectNumber=@p ORDER BY h.Name,h.ComponentHeaderKey'
  $dbt=[Data.DataTable]::new();$reader=$cmd.ExecuteReader();$dbt.Load($reader);$reader.Dispose()
  $formulaPieceTable=$null;$formulaQueryError='';$pieceTable=$null;$databaseMode='Unavailable (no unique DB project)'
  $duplicates=[pscustomobject]@{AffectedTrusses=$null;ExtraPieceRows=$null;Groups=@();Fields=@();Note='No uniquely identified DB project available for piece duplicate detection.'}
  if($pt.Rows.Count -eq 1){
   try{
    $cmd.CommandText='SELECT h.Name AS TrussName,pi.* FROM dbo.Piece pi JOIN dbo.ComponentHeader h ON h.ComponentHeaderKey=pi.ComponentKey WHERE h.ProjectEventKey=@event AND EXISTS (SELECT 1 FROM dbo.ComponentTruss t WHERE t.ComponentKey=pi.ComponentKey) ORDER BY h.Name,pi.ComponentKey,pi.PieceKey'
    $cmd.Parameters.Clear();$null=$cmd.Parameters.Add('@event',[Data.SqlDbType]::UniqueIdentifier);$cmd.Parameters['@event'].Value=$pt.Rows[0].EventKey
    $pieceTable=[Data.DataTable]::new();$reader=$cmd.ExecuteReader();try{$pieceTable.Load($reader)}finally{$reader.Dispose()}
    $duplicates=FindDuplicatePieces $pieceTable
   }catch{$duplicates.Note='Piece duplicate detection unavailable: '+$_.Exception.Message}
   try{
    $cmd.CommandText='SELECT pi.ComponentKey,pi.EngineeringLabel,pi.PlyCount,pi.PickLengthInches,pi.OverallLengthInches,COALESCE(lm.ThicknessNominalUnits,lm.ThicknessInches) AS BdftThickness,COALESCE(lm.WidthNominalUnits,lm.WidthInches) AS BdftWidth FROM dbo.Piece pi JOIN dbo.ComponentHeader h ON h.ComponentHeaderKey=pi.ComponentKey LEFT JOIN dbo.LumberCatalogItem lci ON lci.CatalogItemKey=pi.LumberCatalogItemKey LEFT JOIN dbo.LumberMaterial lm ON lm.LumberMaterialKey=lci.LumberMaterialKey WHERE h.ProjectEventKey=@event AND EXISTS (SELECT 1 FROM dbo.ComponentTruss t WHERE t.ComponentKey=pi.ComponentKey)'
    $cmd.Parameters.Clear();$null=$cmd.Parameters.Add('@event',[Data.SqlDbType]::UniqueIdentifier);$cmd.Parameters['@event'].Value=$pt.Rows[0].EventKey
    $formulaPieceTable=[Data.DataTable]::new();$reader=$cmd.ExecuteReader();try{$formulaPieceTable.Load($reader)}finally{$reader.Dispose()}
   }catch{$formulaPieceTable=$null;$formulaQueryError=$_.Exception.Message}
   try{
    $cmd.CommandText="SELECT PresetValue FROM dbo.PresetData WHERE ProjectKey=@event AND PresetKey=N'Lumber Length'"
    $cmd.Parameters.Clear();$null=$cmd.Parameters.Add('@event',[Data.SqlDbType]::UniqueIdentifier);$cmd.Parameters['@event'].Value=$pt.Rows[0].EventKey
    $presets=[Data.DataTable]::new();$reader=$cmd.ExecuteReader();try{$presets.Load($reader)}finally{$reader.Dispose()}
    $databaseMode=if($presets.Rows.Count -eq 0){'Pick Length (no saved preset; DB procedure default)'}elseif($presets.Rows.Count -ne 1){'Unavailable (multiple Lumber Length presets)'}elseif([string]$presets.Rows[0].PresetValue -in @('Actual Length','Pick Length')){[string]$presets.Rows[0].PresetValue}else{'Unavailable (invalid Lumber Length preset)'}
   }catch{$databaseMode='Unavailable: '+$_.Exception.Message}
  }
  $cmd.Dispose()
  $dbLookup=@{};foreach($db in $dbt.Rows){$dbLookup[$db.Name]=@($dbLookup[$db.Name])+@($db)}
  $projectBody=[Text.StringBuilder]::new();$null=$projectBody.Append('<p><a href="../index.html">Back to all projects</a></p>'+$legend)
  $null=$projectBody.Append((DuplicatePieceDetails $duplicates))
  $pageName=$id+'.html'
  $projectSummaries=[Collections.Generic.List[object]]::new();$archiveIndex=0
  foreach($archive in $groups[$id]){
   $archiveIndex++;$rows=[Collections.Generic.List[object]]::new();$deltas=[Collections.Generic.List[double]]::new();$errors=[Collections.Generic.List[string]]::new()
   $count=@{Synced=0;Differs=0;OnlyFile=0;OnlyDb=0;Marker=0;Older=0;Close=0;Newer=0;Unavailable=0;Ambiguous=0;Skipped=0;Other=0;JsonTrusses=0}
   $archiveWork=if($retainEvidence){Join-Path $run ('evidence\'+$projectIndex+'-'+$archiveIndex)}else{Join-Path $scratchRoot ($projectIndex.ToString()+'-'+$archiveIndex)}
   $null=New-Item -ItemType Directory -Path $archiveWork -Force
   $hash='';$zip=$null;$seen=@{};$bdft=[pscustomobject]@{Total=$null;Note='Archive unavailable';Quantities=@{};Trusses=@();Pieces=@()};$checksums=@();$tdlLookup=@{};$tdlCount=$null;$tdlMatched=0
   try{
    $hash=(Get-FileHash -LiteralPath $archive.FullName -Algorithm SHA256).Hash
    $copy=Join-Path $archiveWork 'source.zip';Copy-Item -LiteralPath $archive.FullName -Destination $copy
    if((Get-FileHash -LiteralPath $copy -Algorithm SHA256).Hash -ne $hash){throw 'Source changed during copy'}
    $zip=[IO.Compression.ZipFile]::OpenRead($copy)
    $tdlEntries=@($zip.Entries | Where-Object {$_.FullName -match '(?i)^Trusses/[^/]+\.tdlTruss$'})
    $tdlCount=$tdlEntries.Count
    foreach($e in $tdlEntries){$n=[IO.Path]::GetFileNameWithoutExtension($e.Name);$tdlLookup[$n]=@($tdlLookup[$n] | Where-Object {$null -ne $_})+@($e)}
    foreach($n in $tdlLookup.Keys){
     if($dbLookup.ContainsKey($n)){$tdlMatched++}else{$count.OnlyFile++}
     if($tdlLookup[$n].Count -gt 1){$count.Ambiguous++}
    }
    $checksums=@(CompareTrussChecksums $zip $dbt.Rows)
    $bdft=EstimateBoardFeet $zip
    $entries=@($zip.Entries | Where-Object {$_.Name -match '(?i)\.json$'} | Sort-Object FullName)
    $count.JsonTrusses=@($entries | Where-Object {$_.FullName -match '(?i)^Trusses/.*\.json$' -or $tdlLookup.ContainsKey([IO.Path]::GetFileNameWithoutExtension($_.Name))}).Count
    $names=@{};foreach($entry in $entries){$name=[IO.Path]::GetFileNameWithoutExtension($entry.Name);$names[$name]=1+$names[$name]}
    $entryIndex=0
    foreach($entry in $entries){
     $entryIndex++;if($entryIndex % 20 -eq 1){Write-Progress -Id 2 -ParentId 1 -Activity $archive.Name -Status "JSON $entryIndex of $($entries.Count)" -PercentComplete ($entryIndex*100/[math]::Max(1,$entries.Count))};$name=[IO.Path]::GetFileNameWithoutExtension($entry.Name);$seen[$name]=$true
     $row=[ordered]@{Truss=$name;Entry=$entry.FullName;Status='';Differences=@();Timestamps=@();FileMatchCode='';DbMatchCode='';SkippedFields=0;Error=''}
     try{
      $dest=Join-Path $archiveWork ('{0:D6}.json' -f $entryIndex)
      [IO.Compression.ZipFileExtensions]::ExtractToFile($entry,$dest)
      $wall=$entry.LastWriteTime.DateTime;[IO.File]::SetLastWriteTime($dest,$wall)
      $j=Get-Content -LiteralPath $dest -Raw | ConvertFrom-Json
      $dbMatches=@($dbLookup[$name] | Where-Object {$null -ne $_})
      $isTruss=$entry.FullName -match '(?i)(^|/)Trusses/' -or $null -ne $j.PSObject.Properties['TrussMatchCode'] -or $dbMatches.Count -gt 0
      if(-not $isTruss){$count.Other++;$row.Status='Other JSON';$rows.Add([pscustomobject]$row);continue}

      if(-not $tdlLookup.ContainsKey($name)){$row.Status='Retained JSON without tdlTruss';$rows.Add([pscustomobject]$row);continue}
      $row.FileMatchCode=[string]$j.TrussMatchCode
      if($names[$name] -gt 1 -or $dbMatches.Count -gt 1 -or $pt.Rows.Count -gt 1){$count.Ambiguous++;$row.Status='Ambiguous name/project';$rows.Add([pscustomobject]$row);continue}
      if($dbMatches.Count -eq 0){$row.Status='Only in csdproj';if($Timestamp){$row.Timestamps=@([pscustomobject]@{Field='ZIP wall time';Value=RawDate $wall})};$rows.Add([pscustomobject]$row);continue}
      $db=$dbMatches[0];$row.DbMatchCode=[string]$db.TrussMatchCode
      foreach($field in $map.Keys){$prop=$j.PSObject.Properties[$field];$b=$db[$map[$field]];if($null -eq $prop -or $null -eq $prop.Value -or $b -is [DBNull]){$row.SkippedFields++;$count.Skipped++;continue};$a=$prop.Value
       $equal=if($a -is [bool]){$a -eq [bool]$b}else{[math]::Abs([double]$a-[double]$b) -le $NumericTolerance}
       if(-not $equal){$row.Differences+= [pscustomobject]@{Field=$field;DatabaseField=$map[$field];csdproj=$a;Database=$b}}
      }
      if($row.SkippedFields -eq $map.Count){$row.Status='No comparable fields';$count.Ambiguous++}elseif($row.Differences.Count){$row.Status='Data differs';$count.Differs++}else{$row.Status='Data synced';$count.Synced++}
      if($row.FileMatchCode -cne $row.DbMatchCode){$count.Marker++}
      if($Timestamp){
foreach($field in @('HeaderModified','ComponentModified','LastModifiedDateTime')){
       $seconds=$null;$note='';$dbvalue=$db[$field]
       try{if($dbvalue -is [DBNull]){throw 'Database timestamp missing'};$seconds=((UtcDate $wall $fileZone)-(UtcDate $dbvalue $dbZone)).TotalSeconds}catch{$note=$_.Exception.Message}
       $row.Timestamps+=[pscustomobject]@{Field=$field;csdprojWallTime=RawDate $wall;DatabaseWallTime=RawDate $dbvalue;Difference=SignedTime $seconds;DifferenceSeconds=$seconds;Note=$note}
       if($field -eq 'LastModifiedDateTime'){if($null -eq $seconds){$count.Unavailable++}else{$deltas.Add($seconds);if([math]::Abs($seconds) -le $ToleranceSeconds){$count.Close++}elseif($seconds -lt 0){$count.Older++}else{$count.Newer++}}}
      }
      }
     }catch{$row.Status='JSON error';$row.Error=$_.Exception.Message;$errors.Add($entry.FullName+': '+$row.Error)}
     $rows.Add([pscustomobject]$row)
    }
    foreach($db in $dbt.Rows){if(-not $tdlLookup.ContainsKey($db.Name)){$count.OnlyDb++;$rows.Add([pscustomobject]@{Truss=$db.Name;Status='Only in database';DbMatchCode=[string]$db.TrussMatchCode;Differences=@();Timestamps=@(if($Timestamp){[pscustomobject]@{Field='Truss modified';Value=RawDate $db.LastModifiedDateTime}})})}}
    foreach($n in $tdlLookup.Keys){
     if(-not $seen.ContainsKey($n)){
      $matchedDb=@($dbLookup[$n] | Where-Object {$null -ne $_})
      if($matchedDb.Count){$count.Ambiguous++;if($Timestamp){$count.Unavailable++};$rows.Add([pscustomobject]@{Truss=$n;Status='tdlTruss in both but JSON unavailable';Differences=@();Timestamps=@();Error=if($Timestamp){'No JSON file available for selected-field or JSON timestamp comparison'}else{'No JSON file available for selected-field comparison'}})}
      else{$rows.Add([pscustomobject]@{Truss=$n;Status='Only in csdproj';Differences=@();Timestamps=@();Error='tdlTruss has no matching DB name and no JSON'})}
     }
    }
    if((Get-FileHash -LiteralPath $archive.FullName -Algorithm SHA256).Hash -ne $hash){throw 'Original archive changed during investigation'}
   }catch{$errors.Add($_.Exception.Message)}finally{if($null -ne $zip){$zip.Dispose()}}
   $min=$null;$max=$null;if($deltas.Count){$m=$deltas | Measure-Object -Minimum -Maximum;$min=$m.Minimum;$max=$m.Maximum}
   $status=if($errors.Count){'Incomplete / errors'}elseif($pt.Rows.Count -eq 0){'Project not in DB'}elseif($count.Ambiguous){'Ambiguous / incomplete'}else{'Compared'}
   if($Timestamp -and $status -eq 'Compared'){
    if($count.Older -gt 0 -and $count.Newer -gt 0){$status='Mixed timestamps (older and newer)'}
    elseif($count.Older -gt 0){$status='Older csdproj timestamps'}
    elseif($count.Newer -gt 0){$status='Newer csdproj timestamps'}
    elseif($count.Close -gt 0){$status='Within tolerance'}
    elseif($count.Unavailable -gt 0){$status='Timestamp unavailable'}
    if($count.Unavailable -gt 0 -and $status -ne 'Timestamp unavailable'){$status+='; some unavailable'}
   }
   if($count.OnlyDb -gt 0 -and $errors.Count -eq 0 -and $pt.Rows.Count -gt 0){
    $hint=if($status -eq 'Compared'){''}else{'; '+$status}
    $status='Missing from csdproj: '+$count.OnlyDb+' DB truss names'+$hint
   }elseif($status -eq 'Compared' -and $count.OnlyFile -gt 0){$status='No matching DB truss names'}
   $dbBdft=$null;if($pt.Rows.Count -eq 1 -and $pt.Rows[0].CachedBoardFeet -isnot [DBNull]){$dbBdft=[double]$pt.Rows[0].CachedBoardFeet}
   if($errors.Count){$bdft.Total=$null;$bdft.Note='Incomplete archive processing; estimate withheld'}
   $reconciliation=ReconcileBoardFeet $bdft $pieceTable $duplicates $dbt.Rows $dbBdft $databaseMode
   $bfDelta=if($null -ne $bdft.Total -and $null -ne $dbBdft){$bdft.Total-[decimal]$dbBdft}else{$null}
   $s=[pscustomobject][ordered]@{Project=$id;csdproj=$archive.Name;Status=$status;JsonTrussCount=$count.JsonTrusses;TdlTrussCount=$tdlCount;DbTrussCount=$dbt.Rows.Count;TdlMatchedNames=$tdlMatched;'Data synced'=$count.Synced;'Data differs'=$count.Differs;ChecksumSame=@($checksums | Where-Object Result -eq 'Same file').Count;ChecksumDifferent=@($checksums | Where-Object Result -eq 'Different file').Count;ChecksumMissing=@($checksums | Where-Object Result -eq 'Missing file').Count;ChecksumUnavailable=if($checksums.Count){@($checksums | Where-Object Result -eq 'Unavailable').Count}else{$dbt.Rows.Count};CsdprojBoardFeet=$bdft.Total;DatabaseBoardFeet=$dbBdft;BoardFeetDifference=$bfDelta;DuplicateAffectedTrusses=$duplicates.AffectedTrusses;DuplicateExtraPieceRows=$duplicates.ExtraPieceRows;'Only in csdproj'=$count.OnlyFile;'Only in database'=$count.OnlyDb;'Review marker differs'=$count.Marker;'Older Timestamp'=$count.Older;'Within tolerance'=$count.Close;'Newer Timestamp'=$count.Newer;Unavailable=$count.Unavailable;'Minimum difference'=SignedTime $min;'Maximum difference'=SignedTime $max;Ambiguous=$count.Ambiguous}
   if(-not $Timestamp){foreach($name in @('Older Timestamp','Within tolerance','Newer Timestamp','Unavailable','Minimum difference','Maximum difference')){$s.PSObject.Properties.Remove($name)}}
   $fileBasisComparison=CompareBoardFeetOnFileBasis $bdft $formulaPieceTable $dbt.Rows $databaseMode
   if($formulaQueryError){$fileBasisComparison.Note+=' SQL material query: '+$formulaQueryError}
   $basisEvidence=DatabaseBoardFeetBasis $pieceTable $dbt.Rows $dbBdft $databaseMode
   $s | Add-Member -NotePropertyName CsdprojBoardFeetBasis -NotePropertyValue (ArchiveBoardFeetBasis $bdft)
   $s | Add-Member -NotePropertyName DatabaseBoardFeetBasis -NotePropertyValue $basisEvidence.CachedBasis
   $s | Add-Member -NotePropertyName DatabasePieceBoardFeetBasis -NotePropertyValue $basisEvidence.PieceBasis
   $s | Add-Member -NotePropertyName BoardFeetBasisNote -NotePropertyValue $basisEvidence.Note
   $s | Add-Member -NotePropertyName DatabaseOnFileBasisBoardFeet -NotePropertyValue $fileBasisComparison.DatabaseTotal
   $s | Add-Member -NotePropertyName SameBasisBoardFeetDifference -NotePropertyValue $fileBasisComparison.Difference
   $s | Add-Member -NotePropertyName SameBasisDifferentTrusses -NotePropertyValue $fileBasisComparison.DifferentTrusses
   $s | Add-Member -NotePropertyName BoardFeetQuantityDifferences -NotePropertyValue $fileBasisComparison.QuantityDifferences
   $s | Add-Member -NotePropertyName BoardFeetSettingsDifferent -NotePropertyValue $fileBasisComparison.SettingsDifferent
   $s | Add-Member -NotePropertyName SameBasisBoardFeetNote -NotePropertyValue $fileBasisComparison.Note
   $dataChecksHealthy=HealthyProject $s
   $s | Add-Member -NotePropertyName DataChecksHealthy -NotePropertyValue $dataChecksHealthy
   $s | Add-Member -NotePropertyName BoardFeetStatus -NotePropertyValue (BoardFeetStatus $s)
   $s | Add-Member -NotePropertyName Healthy -NotePropertyValue ($dataChecksHealthy -and $s.BoardFeetStatus -eq 'BDFT agrees' -and $null -ne $s.DuplicateExtraPieceRows -and $s.DuplicateExtraPieceRows -eq 0)
   if($dataChecksHealthy){$s.Status=if($s.Healthy){'Healthy'}else{'Data matches; '+$s.BoardFeetStatus}}
   $allSummary.Add($s);$projectSummaries.Add($s);$links.Add([pscustomobject]@{Summary=$s;Page=$pageName})
   $null=$projectBody.Append('<h2>'+(Enc $archive.Name)+'</h2><div class="scroll">'+(SummaryTable @($s))+'</div><p><b>BDFT:</b> '+(Enc $bdft.Note)+'</p><p>Other JSON: '+$count.Other+'; skipped field comparisons: '+$count.Skipped+'.</p><details><summary>Source evidence</summary><p>'+(Enc $archive.FullName)+'<br>SHA-256: '+$hash+'</p></details>')
   $null=$projectBody.Append((BoardFeetDetails $bdft $reconciliation $dbBdft $basisEvidence $fileBasisComparison))
   $s | Add-Member -NotePropertyName DatabaseLengthMode -NotePropertyValue $databaseMode
   $s | Add-Member -NotePropertyName DatabasePieceBoardFeet -NotePropertyValue $reconciliation.DatabasePieceSum
   $s | Add-Member -NotePropertyName DuplicateBoardFeet -NotePropertyValue $reconciliation.DuplicateContribution
   $s | Add-Member -NotePropertyName DatabaseWithoutDuplicateBoardFeet -NotePropertyValue $reconciliation.DatabaseWithoutDuplicates
   $s | Add-Member -NotePropertyName DatabaseCachedLessDuplicateBoardFeet -NotePropertyValue $reconciliation.CachedLessDuplicates
   $s | Add-Member -NotePropertyName BoardFeetNote -NotePropertyValue $bdft.Note
   foreach($errorText in $errors){$null=$projectBody.Append('<p class="notice">'+(Enc $errorText)+'</p>')}
   $null=$projectBody.Append('<p class="notice"><b>Only data differences, checksum differences or missing files, missing trusses, ambiguities and errors are listed below.</b> Data-synced trusses with matching checksums remain counted above. '+$(if($Timestamp){'Timestamp-only and review-marker-only'}else{'Review-marker-only'})+' differences are excluded from this detail list; the full JSON inventory is retained only with -Evidence True.</p>')
   $checksumByName=@{};foreach($check in $checksums){$checksumByName[$check.Truss]=@($checksumByName[$check.Truss] | Where-Object {$null -ne $_})+@($check)}
   foreach($row in $rows){$row | Add-Member -NotePropertyName FileChecksum -NotePropertyValue @($checksumByName[$row.Truss] | Where-Object {$null -ne $_})}
   $checksumExceptions=@($checksums | Where-Object Result -ne 'Same file')
   $null=$projectBody.Append('<details><summary><b>Truss file checksum exceptions ('+$checksumExceptions.Count+')</b></summary><p>MD5 of individual uncompressed canonical .tdlTruss files versus ComponentHeader.TrussFileCheckSum. Matches are counted in the summary. This is separate from JSON field comparison.</p>'+(Grid $checksumExceptions)+'</details>')
   foreach($row in $rows){if($row.Status -in @('Other JSON','Retained JSON without tdlTruss') -or ($row.Status -eq 'Data synced' -and @($row.FileChecksum | Where-Object Result -ne 'Same file').Count -eq 0)){continue};$null=$projectBody.Append('<details class="'+$(if($row.Status -eq 'Only in database'){'error'}else{''})+'"><summary><b>'+(Enc $row.Truss)+'</b> &mdash; '+(Enc $row.Status)+'</summary><p>'+(Enc $row.Entry)+'</p>');if($row.FileChecksum.Count){$null=$projectBody.Append((Grid $row.FileChecksum));if($row.Differences.Count -gt 0 -and @($row.FileChecksum | Where-Object Result -eq 'Same file').Count -gt 0){$null=$projectBody.Append('<p class="notice"><b>File checksum agrees, but checked JSON/DB fields differ.</b> The stored file checksum does not establish consistency of related database data.</p>')}};if($row.Differences.Count){$null=$projectBody.Append((Grid $row.Differences))};$null=$projectBody.Append('<p>CSEngineer TrussMatchCode<br>csdproj: <code>'+(Enc $row.FileMatchCode)+'</code><br>Database: <code>'+(Enc $row.DbMatchCode)+'</code></p>');if($Timestamp -and $row.Timestamps.Count){$null=$projectBody.Append((Grid $row.Timestamps))};$null=$projectBody.Append('<p>'+(Enc $row.Error)+'</p></details>')}
   if($retainEvidence){[pscustomobject]@{ArchiveCalculation=$bdft;DatabaseReconciliation=$reconciliation;SameBasisComparison=$fileBasisComparison} | ConvertTo-Json -Depth 9 | Set-Content -LiteralPath (Join-Path $archiveWork 'board-feet.json') -Encoding UTF8;$duplicates | ConvertTo-Json -Depth 9 | Set-Content -LiteralPath (Join-Path $archiveWork 'duplicate-pieces.json') -Encoding UTF8;$checksums | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $archiveWork 'checksums.json') -Encoding UTF8;$rows | ConvertTo-Json -Depth 9 | Set-Content -LiteralPath (Join-Path $archiveWork 'comparison.json') -Encoding UTF8}
   Write-Progress -Id 2 -Activity $archive.Name -Completed
   WriteConsoleStatsRow $consoleStatsColumns $s $archiveIndex $groups[$id].Count ([bool]$Timestamp)
  }
  Page $id $projectBody.ToString() (Join-Path $pages $pageName)
 }
}finally{
 $conn.Dispose()
 if($null -ne $scratchRoot -and (Test-Path -LiteralPath $scratchRoot)){
  $resolvedScratch=(Get-Item -LiteralPath $scratchRoot).FullName
  $resolvedTemp=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
  if(-not $resolvedScratch.StartsWith($resolvedTemp,[StringComparison]::OrdinalIgnoreCase) -or -not ([IO.Path]::GetFileName($resolvedScratch)).StartsWith('CsdprojComparison-')){throw 'Unexpected temporary cleanup path'}
  Remove-Item -LiteralPath $resolvedScratch -Recurse -Force
 }
}
$table=[Text.StringBuilder]::new()
$numericColumns=@('Data synced','Data differs','Only in csdproj','Only in database','Review marker differs','Ambiguous','DuplicateAffectedTrusses','DuplicateExtraPieceRows');if($Timestamp){$numericColumns+=@('Older Timestamp','Within tolerance','Newer Timestamp','Unavailable')}
$normalLinks=[Collections.Generic.List[object]]::new()
$emptyLinks=[Collections.Generic.List[object]]::new()
$missingDbLinks=[Collections.Generic.List[object]]::new()
foreach($link in $links){
 $isEmpty=$link.Summary.Status -eq 'Compared'
 foreach($col in $numericColumns){if($link.Summary.$col -ne 0){$isEmpty=$false}}
 if($link.Summary.Status -eq 'Project not in DB'){$missingDbLinks.Add($link)}elseif($isEmpty){$emptyLinks.Add($link)}else{$normalLinks.Add($link)}
}
$columns=if($allSummary.Count){@($allSummary[0].PSObject.Properties.Name)}else{@()}
$null=$table.Append('<div class="scroll overview-table">'+(SummaryHeader))
foreach($section in @('Results','Project not in DB','Empty')){
 $sectionLinks=if($section -eq 'Results'){$normalLinks}elseif($section -eq 'Project not in DB'){$missingDbLinks}else{$emptyLinks}
 $summaryColumnCount=if($Timestamp){27}else{21}
 $null=$table.Append('<tr class="section-divider"><td colspan="'+$summaryColumnCount+'">'+$section+' ('+$sectionLinks.Count+')</td></tr>')
 foreach($link in $sectionLinks){$null=$table.Append((SummaryRow $link.Summary ('projects/'+[uri]::EscapeDataString($link.Page))))}
}
$null=$table.Append('</tbody></table></div>')
$context='{0} projects; {1} csdproj archives. {2} / {3}.' -f $selectedIds.Count,$allSummary.Count,$Server,$Database
$indexLegend=ReportLegend $FileTimeZone $DatabaseTimeZone $ToleranceSeconds $NumericTolerance ([bool]$Timestamp) $context
Page 'CSDirector: Database and Project file Sync Report' ($indexLegend+$table.ToString()) (Join-Path $out 'index.html') -Overview
$allSummary | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $run 'summary.json') -Encoding UTF8
[pscustomobject]@{Run=$runId;Source=$root;Server=$Server;Database=$Database;FileTimeZone=$FileTimeZone;DatabaseTimeZone=$DatabaseTimeZone;ToleranceSeconds=$ToleranceSeconds;NumericTolerance=$NumericTolerance;Fields=$map;RetainEvidence=$retainEvidence;Timestamp=[bool]$Timestamp} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $run 'settings.json') -Encoding UTF8
Write-Progress -Id 2 -Activity 'Files' -Completed
Write-Progress -Id 1 -Activity 'Projects' -Completed
$reportPath=Join-Path $out 'index.html'
$reportUri=([System.Uri]$reportPath).AbsoluteUri
$escape=[char]27
# OSC 8 enables a clickable hyperlink in Windows Terminal and compatible terminals.
# The visible file URI remains usable in terminals without hyperlink support.
Write-Host ("Report: ${escape}]8;;${reportUri}${escape}\${reportUri}${escape}]8;;${escape}\")








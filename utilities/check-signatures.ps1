# Check-Signatures.ps1
#
# Recursively checks Authenticode signatures for DLL and EXE files.
#
# Examples:
#   .\Check-Signatures.ps1
#   .\Check-Signatures.ps1 -Path "C:\SST"
#   .\Check-Signatures.ps1 "C:\Program Files\MyApplication"

[CmdletBinding()]
param (
    [Parameter(
        Position = 0,
        Mandatory = $false,
        HelpMessage = "Directory containing the DLL and EXE files to scan."
    )]
    [ValidateNotNullOrEmpty()]
    [string]$Path = (Get-Location).Path
)

# Validate and resolve the supplied directory.
if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
    Write-Error "The directory does not exist or is not accessible: $Path"
    exit 1
}

$ScanPath = (Resolve-Path -LiteralPath $Path).Path
$ReportPath = Join-Path $ScanPath "SignatureReport.csv"

Write-Host ""
Write-Host "Scanning: $ScanPath" -ForegroundColor Cyan

$files = @(
    Get-ChildItem -LiteralPath $ScanPath -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Extension -in @(".dll", ".exe")
        }
)

if ($files.Count -eq 0) {
    Write-Warning "No DLL or EXE files were found under: $ScanPath"
    exit 0
}

$results = foreach ($file in $files) {
    try {
        $signature = Get-AuthenticodeSignature -LiteralPath $file.FullName

        [PSCustomObject]@{
            FileName         = $file.Name
            FullPath         = $file.FullName
            Status           = $signature.Status
            StatusMessage    = $signature.StatusMessage
            SignatureType    = $signature.SignatureType
            Publisher        = if ($signature.SignerCertificate) {
                $signature.SignerCertificate.Subject
            }
            else {
                ""
            }
            Issuer           = if ($signature.SignerCertificate) {
                $signature.SignerCertificate.Issuer
            }
            else {
                ""
            }
            CertificateStart = if ($signature.SignerCertificate) {
                $signature.SignerCertificate.NotBefore
            }
            else {
                $null
            }
            CertificateEnd   = if ($signature.SignerCertificate) {
                $signature.SignerCertificate.NotAfter
            }
            else {
                $null
            }
            Thumbprint       = if ($signature.SignerCertificate) {
                $signature.SignerCertificate.Thumbprint
            }
            else {
                ""
            }
        }
    }
    catch {
        [PSCustomObject]@{
            FileName         = $file.Name
            FullPath         = $file.FullName
            Status           = "CheckFailed"
            StatusMessage    = $_.Exception.Message
            SignatureType    = ""
            Publisher        = ""
            Issuer           = ""
            CertificateStart = $null
            CertificateEnd   = $null
            Thumbprint       = ""
        }
    }
}

$results = @($results | Sort-Object Status, FullPath)

$results |
    Format-Table Status, SignatureType, FileName, Publisher -AutoSize

$results |
    Export-Csv -LiteralPath $ReportPath -NoTypeInformation -Encoding UTF8

$validCount = @(
    $results | Where-Object Status -eq "Valid"
).Count

$unsignedCount = @(
    $results | Where-Object Status -eq "NotSigned"
).Count

$problemCount = @(
    $results | Where-Object {
        $_.Status -notin @("Valid", "NotSigned")
    }
).Count

Write-Host ""
Write-Host "Scan complete" -ForegroundColor Cyan
Write-Host "  Total files: $($results.Count)"
Write-Host "  Valid:       $validCount" -ForegroundColor Green
Write-Host "  Not signed:  $unsignedCount" -ForegroundColor Yellow
Write-Host "  Other issue: $problemCount" -ForegroundColor Red
Write-Host ""
Write-Host "Report saved to: $ReportPath" -ForegroundColor Cyan
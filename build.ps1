# build.ps1 - Cross-Platform Multi-Target Publish Script
$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $env:TEMP) { $env:TEMP = "$scriptDir\.tmp" }
if (-not $env:TMP) { $env:TMP = "$scriptDir\.tmp" }
if (-not $env:NUGET_PACKAGES) { $env:NUGET_PACKAGES = "$scriptDir\.nuget\packages" }

$rids = @("win-x86", "win-x64", "linux-x64", "linux-arm", "linux-arm64")

foreach ($rid in $rids) {
    Write-Host "Building CrossReportRunner for $rid..." -ForegroundColor Cyan
    dotnet publish "$scriptDir/src/CrossReportRunner/CrossReportRunner.csproj" `
        -c Release `
        -r $rid `
        --self-contained true `
        -p:PublishSingleFile=true `
        -p:IncludeNativeLibrariesForSelfExtract=true `
        -p:EnableCompressionInSingleFile=true `
        -o "$scriptDir/bin/$rid"
}
Write-Host "All targets published successfully." -ForegroundColor Green

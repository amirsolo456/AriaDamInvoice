$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$projectPath = Join-Path $repoRoot "AriaDamInvoice.App\AriaDamInvoice.App.csproj"
$publishPath = Join-Path $repoRoot "artifacts\AriaDamInvoice-Portable"
$expectedExe = Join-Path $publishPath "AriaDamInvoice.exe"

if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
    throw "The .NET 10 SDK is required on the development computer to publish this application."
}

Write-Host "Publishing AriaDamInvoice as a self-contained Windows x64 application..." -ForegroundColor Cyan

$publishArgs = @(
    "publish"
    $projectPath
    "--configuration", "Release"
    "--runtime", "win-x64"
    "--self-contained", "true"
    "-p:PublishSingleFile=false"
    "-p:IncludeNativeLibrariesForSelfExtract=true"
    "--output", $publishPath
)
& dotnet @publishArgs

if ($LASTEXITCODE -ne 0) {
    throw "dotnet publish failed. Review the build output above."
}

if (-not (Test-Path $expectedExe)) {
    throw "Publish completed but AriaDamInvoice.exe was not found at: $expectedExe"
}

Write-Host ""
Write-Host "Portable folder is ready:" -ForegroundColor Green
Write-Host $publishPath -ForegroundColor Yellow
Write-Host ""
Write-Host "Copy the ENTIRE folder to the destination computer." -ForegroundColor Green
Write-Host "Also copy the original Excel template workbook separately."
Write-Host "Microsoft Excel desktop and the B Nazanin / B Yekan fonts must be installed on the destination computer."

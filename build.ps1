# Build gongwen-paiban as a self-contained single-file executable.
# Usage: .\build.ps1 [-Rid win-x64]

param([string]$Rid = "win-x64")

$ErrorActionPreference = 'Stop'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$outDir = Join-Path $scriptDir "dist/$Rid"

Write-Host "==> Building gongwen-paiban for $Rid -> $outDir"

dotnet publish (Join-Path $scriptDir "src/MiniMaxAIDocx.Cli/MiniMaxAIDocx.Cli.csproj") `
    -c Release -r $Rid --self-contained true `
    -p:PublishSingleFile=true `
    -p:IncludeNativeLibrariesForSelfExtract=true `
    -p:EnableCompressionInSingleFile=true `
    -o $outDir
if ($LASTEXITCODE -ne 0) { throw "dotnet publish failed" }

if ($Rid -eq "win-x64") {
    Move-Item -Force (Join-Path $outDir "MiniMaxAIDocx.Cli.exe") (Join-Path $outDir "gongwen-paiban.exe")
    Copy-Item (Join-Path $scriptDir "launcher/format.bat") (Join-Path $outDir "format.bat")
    Copy-Item (Join-Path $scriptDir "launcher/format.ps1") (Join-Path $outDir "format.ps1")
} else {
    $bin = Join-Path $outDir "MiniMaxAIDocx.Cli"
    if (Test-Path $bin) { Move-Item -Force $bin (Join-Path $outDir "gongwen-paiban") }
}

Get-ChildItem "$outDir/*.pdb" -ErrorAction SilentlyContinue | Remove-Item -Force

Write-Host "==> Done."
Get-ChildItem $outDir

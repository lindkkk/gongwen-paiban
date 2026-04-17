# 公文排版工具安装脚本 (Windows)

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$DotnetDir = Join-Path $ScriptDir "dotnet"

Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "  公文排版工具安装程序" -ForegroundColor Cyan
Write-Host "========================================="
Write-Host ""

# 检查 .NET SDK
try {
    $dotnetVersion = dotnet --version
    Write-Host ".NET SDK version: $dotnetVersion" -ForegroundColor Green
} catch {
    Write-Host "Error: .NET SDK is not installed." -ForegroundColor Red
    Write-Host "Please install .NET 8.0 SDK from: https://dotnet.microsoft.com/download" -ForegroundColor Yellow
    exit 1
}

# 还原项目
Write-Host ""
Write-Host "Restoring NuGet packages..." -ForegroundColor Yellow
Set-Location $DotnetDir
dotnet restore

# 构建项目
Write-Host ""
Write-Host "Building project..." -ForegroundColor Yellow
dotnet build --configuration Release --no-restore

Write-Host ""
Write-Host "=========================================" -ForegroundColor Green
Write-Host "  安装完成!" -ForegroundColor Green
Write-Host "========================================="
Write-Host ""
Write-Host "Usage:" -ForegroundColor White
Write-Host '  dotnet run --project $DotnetDir\MiniMaxAIDocx.Cli -- format input.docx -o output.docx' -ForegroundColor Gray
Write-Host ""

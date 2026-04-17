#!/bin/bash
# 公文排版工具安装脚本

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
DOTNET_DIR="$SCRIPT_DIR/dotnet"

echo "========================================="
echo "  公文排版工具安装程序"
echo "========================================="
echo ""

# 检查 .NET SDK
if ! command -v dotnet &> /dev/null; then
    echo "Error: .NET SDK is not installed."
    echo "Please install .NET 8.0 SDK from: https://dotnet.microsoft.com/download"
    exit 1
fi

echo "Checking .NET SDK version..."
dotnet --version

# 还原项目
echo ""
echo "Restoring NuGet packages..."
cd "$DOTNET_DIR"
dotnet restore

# 构建项目
echo ""
echo "Building project..."
dotnet build --configuration Release --no-restore

echo ""
echo "========================================="
echo "  安装完成!"
echo "========================================="
echo ""
echo "Usage:"
echo "  dotnet run --project $DOTNET_DIR/MiniMaxAIDocx.Cli -- format input.docx -o output.docx"
echo ""

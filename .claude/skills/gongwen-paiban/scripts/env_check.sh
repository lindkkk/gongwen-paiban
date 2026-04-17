#!/bin/bash
# 环境检查脚本

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
DOTNET_DIR="$SCRIPT_DIR/dotnet"

NOT_READY=0

echo "Checking environment..."

# 检查 .NET SDK
if ! command -v dotnet &> /dev/null; then
    echo "[FAIL] .NET SDK not found"
    NOT_READY=1
else
    echo "[PASS] .NET SDK: $(dotnet --version)"
fi

# 检查项目文件
if [ ! -f "$DOTNET_DIR/MiniMaxAIDocx.Core/MiniMaxAIDocx.Core.csproj" ]; then
    echo "[FAIL] Project files not found"
    NOT_READY=1
else
    echo "[PASS] Project files exist"
fi

# 检查 NuGet 包是否已还原
if [ -d "$DOTNET_DIR/obj" ]; then
    echo "[PASS] Build artifacts exist"
else
    echo "[WARN] Run setup.sh first to restore packages"
    NOT_READY=1
fi

if [ $NOT_READY -eq 1 ]; then
    echo ""
    echo "NOT READY - Please run setup.sh first"
    exit 1
fi

echo ""
echo "READY"
exit 0

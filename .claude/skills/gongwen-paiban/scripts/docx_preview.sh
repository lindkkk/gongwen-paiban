#!/bin/bash
# DOCX文档预览脚本

if [ $# -eq 0 ]; then
    echo "Usage: $0 <document.docx>"
    exit 1
fi

INPUT_FILE="$1"

if [ ! -f "$INPUT_FILE ]; then
    echo "Error: File not found: $INPUT_FILE"
    exit 1
fi

echo "Previewing: $INPUT_FILE"
echo "=========================================="

# 使用CLI预览
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
CLI="$SCRIPT_DIR/dotnet/MiniMaxAIDocx.Cli"

if [ -f "$CLI" ] || [ -f "$CLI.exe" ]; then
    dotnet run --project "$CLI" -- preview "$INPUT_FILE"
else
    echo "Error: CLI not built. Run setup.sh first."
    exit 1
fi

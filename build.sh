#!/usr/bin/env bash
# Build gongwen-paiban as a self-contained single-file executable.
# Usage: ./build.sh [win-x64|linux-x64|osx-x64|osx-arm64]   default: win-x64

set -euo pipefail
RID="${1:-win-x64}"

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
OUT_DIR="$SCRIPT_DIR/dist/$RID"

echo "==> Building gongwen-paiban for $RID → $OUT_DIR"

dotnet publish "$SCRIPT_DIR/src/MiniMaxAIDocx.Cli/MiniMaxAIDocx.Cli.csproj" \
    -c Release -r "$RID" --self-contained true \
    -p:PublishSingleFile=true \
    -p:IncludeNativeLibrariesForSelfExtract=true \
    -p:EnableCompressionInSingleFile=true \
    -o "$OUT_DIR"

# Rename to a friendly binary name
if [ "$RID" = "win-x64" ]; then
    mv -f "$OUT_DIR/MiniMaxAIDocx.Cli.exe" "$OUT_DIR/gongwen-paiban.exe"
    # Copy launcher scripts so the win-x64 dist is a complete drag-drop package
    cp "$SCRIPT_DIR/launcher/format.bat" "$OUT_DIR/"
    cp "$SCRIPT_DIR/launcher/format.ps1" "$OUT_DIR/"
else
    mv -f "$OUT_DIR/MiniMaxAIDocx.Cli" "$OUT_DIR/gongwen-paiban" 2>/dev/null || true
    chmod +x "$OUT_DIR/gongwen-paiban" 2>/dev/null || true
fi

rm -f "$OUT_DIR"/*.pdb

echo "==> Done."
ls -lah "$OUT_DIR"

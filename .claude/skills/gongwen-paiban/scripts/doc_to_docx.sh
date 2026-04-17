#!/bin/bash
# DOC转DOCX转换脚本
# 注意: 需要安装LibreOffice或Microsoft Word

if [ $# -lt 1 ]; then
    echo "Usage: $0 <input.doc> [output_dir]"
    exit 1
fi

INPUT_FILE="$1"
OUTPUT_DIR="${2:-.}"

if [ ! -f "$INPUT_FILE" ]; then
    echo "Error: File not found: $INPUT_FILE"
    exit 1
fi

FILENAME=$(basename "$INPUT_FILE")
BASENAME="${FILENAME%.*}"
OUTPUT_FILE="$OUTPUT_DIR/$BASENAME.docx"

echo "Converting: $INPUT_FILE"
echo "Output: $OUTPUT_FILE"

# 尝试使用LibreOffice进行转换
if command -v libreoffice &> /dev/null; then
    libreoffice --headless --convert-to docx --outdir "$OUTPUT_DIR" "$INPUT_FILE"
    echo "Conversion complete!"
elif command -v soffice &> /dev/null; then
    soffice --headless --convert-to docx --outdir "$OUTPUT_DIR" "$INPUT_FILE"
    echo "Conversion complete!"
else
    echo "Error: Neither LibreOffice nor soffice found."
    echo "Please convert the .doc file to .docx using Microsoft Word manually,"
    echo "then use the format command."
    exit 1
fi

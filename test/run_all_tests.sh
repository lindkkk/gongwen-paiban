#!/bin/bash
# 统一跑一遍所有测试样本
set -e
export PATH="$HOME/.dotnet:$PATH"

TEST_DIR="/home/asus11700f/下载/doc-skill/test"
DOTNET_DIR="/home/asus11700f/下载/doc-skill/.claude/skills/gongwen-paiban/scripts/dotnet"

cd "$DOTNET_DIR"
dotnet build MiniMaxAIDocx.Cli/MiniMaxAIDocx.Cli.csproj --configuration Release 2>&1 | tail -3

declare -a CASES=(
    "raw_test|expect_cover=1|min_h1=3|min_h2=3|min_h3=3|min_body=20|expect_toc_section=0"
    "raw_long_toc|expect_cover=1|min_h1=5|min_h2=5|min_h3=5|min_body=20|expect_toc_section=1"
    "raw_plain|expect_cover=0|min_h1=0|min_h2=0|min_h3=0|min_body=50|expect_toc_section=0"
    "raw_with_styles|expect_cover=1|min_h1=3|min_h2=3|min_h3=3|min_body=20|expect_toc_section=0"
    "raw_wps_polluted|expect_cover=0|min_h1=0|min_h2=5|min_h3=0|min_body=10|expect_toc_section=0"
    "raw_auto_multilevel|expect_cover=0|min_h1=3|min_h2=2|min_h3=2|min_body=5|expect_toc_section=0"
    "raw_decimal_multilevel|expect_cover=0|min_h1=3|min_h2=3|min_h3=3|min_body=10|expect_toc_section=0"
)

OVERALL_OK=1
for line in "${CASES[@]}"; do
    name="${line%%|*}"
    rest="${line#*|}"
    echo ""
    echo "===================================================="
    echo "==> 样本: $name"
    echo "===================================================="
    # 解析参数
    expect_cover=0; expect_toc_section=0
    min_h1=0; min_h2=0; min_h3=0; min_body=10
    IFS='|' read -ra KVS <<< "$rest"
    for kv in "${KVS[@]}"; do
        key="${kv%%=*}"; val="${kv#*=}"
        declare "$key=$val"
    done

    in_file="$TEST_DIR/${name}.docx"
    out_file="$TEST_DIR/formatted_${name}.docx"

    dotnet run --project MiniMaxAIDocx.Cli --configuration Release --no-build -- \
        format "$in_file" -o "$out_file" 2>&1 | tail -2

    flags=()
    [ "$expect_cover" = "1" ] && flags+=(--expect-cover)
    [ "$expect_toc_section" = "1" ] && flags+=(--expect-toc-section)
    flags+=(--min-h1 "$min_h1" --min-h2 "$min_h2" --min-h3 "$min_h3" --min-body "$min_body")

    if ! python3 "$TEST_DIR/verify.py" "$out_file" "${flags[@]}" ; then
        OVERALL_OK=0
        echo "❌ $name 校验失败"
    fi
done

echo ""
echo "===================================================="
if [ "$OVERALL_OK" = "1" ]; then
    echo "✅ 全部样本通过"
    exit 0
else
    echo "❌ 存在样本失败"
    exit 1
fi

# 开发指南

面向想改源码、构建二进制或扩展功能的人。

## 环境

- **.NET 8 SDK**（必需）——[Microsoft 下载页](https://dotnet.microsoft.com/download/dotnet/8.0)，Windows / Linux / macOS 都有一键安装包
- **Python 3.8+**（仅跑测试用）+ `python-docx` + `lxml`
- **可选**：LibreOffice（把 `.doc` 转 `.docx` 用）

Linux 无 sudo 装 .NET 8：

```bash
curl -fsSL https://dot.net/v1/dotnet-install.sh -o /tmp/dotnet-install.sh
bash /tmp/dotnet-install.sh --channel 8.0 --install-dir $HOME/.dotnet
export PATH="$HOME/.dotnet:$PATH"
dotnet --version   # 确认 8.0.xxx
```

## 构建

### 一键构建（推荐）

```bash
# Linux / macOS：
./build.sh win-x64            # 输出 dist/win-x64/gongwen-paiban.exe + 两个 launcher
./build.sh linux-x64
./build.sh osx-arm64

# Windows PowerShell：
.\build.ps1 win-x64
```

产物约 36 MB，零依赖双击即用（Windows 7 SP1 及以上）。

### 开发时的快速构建

```bash
dotnet build src/MiniMaxAIDocx.Cli/MiniMaxAIDocx.Cli.csproj -c Release
```

输出在 `src/MiniMaxAIDocx.Cli/bin/Release/net8.0/`，运行需要装 .NET 运行时。

## 跑测试

```bash
cd test
python3 gen_test_doc.py           # 生成各种 fixture
python3 gen_long_toc.py
python3 gen_plain.py
python3 gen_with_styles.py
python3 gen_wps_polluted.py
python3 gen_auto_multilevel.py
python3 gen_decimal_multilevel.py

bash run_all_tests.sh             # 一次性跑全部场景
```

测试会：

1. 调构建好的 exe 格式化每份 `raw_*.docx`
2. 用 `verify.py` 对每段逐字段断言（字体 / 字号 / 粗体 / 行距 / 段前段后 / 首行缩进 / 对齐 / Section 页码类型 / 无 theme 字体残留 / 无 numPr 残留）
3. 任意断言不通过立刻标红退出

7 个当前场景：

| 样本 | 覆盖点 |
|---|---|
| `raw_test` | 中文编号 + 短 TOC + 50+ 页长文 |
| `raw_long_toc` | 目录 ≥2 页，测 Roman 页码分支 |
| `raw_plain` | 无结构纯正文 |
| `raw_with_styles` | 用户已用 Word 内置 Heading 1/2/3 |
| `raw_wps_polluted` | 模拟 WPS 塞的主题字体 + numPr + 脏 pPr |
| `raw_auto_multilevel` | H1/H2/H3 都用 numPr 自动编号，不同 ilvl |
| `raw_decimal_multilevel` | 1 / 1.1 / 1.1.1 手工十进制多级 |

## 常见改动指引

### 改字体 / 字号 / 行距

全在 `GongWenFormatter.cs` 顶部常量区。改一行就行。

```csharp
public const string FontH1 = "黑体";         // 想改就改值
public const string FontSizeSanhao = "32";   // 16pt × 2，想换 15pt 就写 "30"
public const int FirstLineChars = 200;       // 2 字符 = 200
```

### 加新样式（例如"四级标题"）

1. `GongWenFormatter.cs`：加 `StyleIdHeading4` 常量、`FontH4` 字体常量
2. `CreateGongWenStyles` 里 `AddOrReplaceStyle` 注册样式定义
3. 在 `ParaRole` 枚举加 `H4`
4. `ClassifyHeadingByText` / `ClassifyRemaining` 增加规则
5. `ApplyFormatting` switch 加 `case H4`，写对应 `ApplyHeading4Style`
6. `test/verify.py` 的 `SPECS` / `STYLEID_TO_ROLE` 同步加
7. 加一个测试 fixture `gen_with_h4.py`

### 改 TOC / 页码行为

主要在 `SetupSections` 里。封面、TOC、正文三个 section 的 `BuildSectionProperties` 调用是入口。

例：想让**目录不管几页都用 Roman**，把 `hasTocSection` 判断从 `TocPageCount >= 2` 改成 `TocStartIndex.HasValue`。

例：想让**封面显示页码**（少见但规范有时要），在封面 section 的 `BuildSectionProperties` 调用里给 `footerRefId` 换成 arabic footer，把 `pageFormat` 设为 `Decimal`。

### 加新的用户 CLI 参数

1. `FormatOptions.cs` 加字段
2. `Program.cs` 的 `FormatDocument` 里加 `else if (a == "--xxx") { options.XXX = next(); }`
3. `GongWenFormatter.FormatDocument` 里使用 `_options.XXX`

### 同步到 PowerShell UI

如果新参数要从对话框收集：

1. `dist/win-x64/format.ps1`：加新控件、读取文本框值
2. 把值塞到 `$exeArgs` 数组里：`if ($newVal) { $exeArgs += @("--xxx", $newVal) }`

## 当心几个坑

1. **不要用 `ProcessStartInfo.ArgumentList`**——PS 5.1 跑在 .NET Framework 4.x 上没这属性。用调用运算符 `& $exe @args`。

2. **bat 永远保持纯 ASCII + CRLF 换行**——中文 Windows 的 cmd 默认 GBK 码页，bat 里有非 ASCII 字节会被当 GBK 解释成乱码。UI 文字都放 ps1。

3. **ps1 永远保持 UTF-8 with BOM**——PS 5.1 没 BOM 按系统码页读 .ps1。我用 Python 脚本生成文件确保 BOM，别用 `notepad → 另存为` 会丢 BOM。

4. **bat 不要对 powershell 调用整段重定向** `1>>"%LOG%"`——那会在整个 PS 执行期间独占写锁，ps1 的 `Add-Content` 会静默失败。让 ps1 自己写日志。

5. **rPr 不要有 `*Theme` 属性**——theme 字体优先级高于显式字体，写了就会被主题覆盖。`RebuildRPr` 整个重建 rPr 就是为此。

6. **OpenXML schema 顺序不能乱**——`pStyle` 必须是 pPr 的第一个子元素，`rFonts` 通常应在 rPr 最前。`RebuildPPr/RebuildRPr` 按正确顺序 `AppendChild` 即可。

7. **`Indentation.FirstLineChars` 是 `Int32Value` 不是字符串**——早期我用字符串常量编译不过。现在常量定义为 `int`。

## 调试技巧

- 单元测 MarkerPatternInferrer：`dotnet run -- test-marker "1.1" "1.1 标题" "1.1.1 标题"`
- 纯命令行验证 exe：绕开 bat / ps1，直接 `gongwen-paiban.exe format in.docx -o out.docx`
- 看生成的 docx 内部 XML：`unzip out.docx -d /tmp/out && cat /tmp/out/word/document.xml`
- PS 脚本调试：在 ps1 里多加 `Log "xxx"`，看 `paiban-log.txt`

## 扩展方向（今后想做但还没做的）

- 目录段落按规范排版（当前只保持原样不动）
- 手打脚注段落识别（当前仅认 Word 原生 FootnotesPart）
- 表格内段落也格式化（当前跳过）
- SDT 包裹的 TOC 支持（Word 自动生成的 TOC）
- `.doc` 直接读取（用 NPOI 或调 LibreOffice 转换）
- 双面打印空白页避免加页码
- 代码签名 exe 消除 SmartScreen 警告

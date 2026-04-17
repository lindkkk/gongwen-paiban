# Contributing · 贡献指南

[English](#english) · [中文](#中文)

---

## 中文

感谢愿意贡献！这是一个**单人维护、但欢迎改进**的项目。下面列些注意点。

### 开发环境

- **.NET 8 SDK**（必需）
- **Python 3.8+** + `python-docx` + `lxml`（跑测试用）
- **PowerShell 5.1+**（Windows 自带 / 或 `pwsh` 7.x）

Linux 无 sudo 装 .NET：
```bash
curl -fsSL https://dot.net/v1/dotnet-install.sh | bash /dev/stdin --channel 8.0 --install-dir $HOME/.dotnet
export PATH="$HOME/.dotnet:$PATH"
```

### 目录结构

```
src/                                      核心 C# 源码
  ├── MiniMaxAIDocx.Core/
  │   ├── GongWenFormatter.cs             主入口：分类器 / 样式应用 / section 设置
  │   ├── NumberingResolver.cs            把 numPr 自动编号固化成文字
  │   ├── MarkerPatternInferrer.cs        用户样例 → 宽松 regex
  │   ├── FormatOptions.cs / StyleSpec.cs 配置选项 + 单角色规格
  └── MiniMaxAIDocx.Cli/
      └── Program.cs                      CLI 入口
launcher/                                 Windows 启动器（拖拽 + 对话框）
  ├── format.bat                          纯 ASCII + CRLF（别用编辑器另存）
  └── format.ps1                          UTF-8 BOM + CRLF
test/                                     跑测脚本
docs/                                     设计文档
.github/workflows/                        GitHub Actions
```

### 流程

1. Fork → clone → 改代码
2. 本地构建：`./build.sh win-x64`
3. 跑测试：`cd test && bash run_all_tests.sh`（7 个 fixture 必须全绿）
4. 提 PR，描述**问题现象 + 根因分析 + 修复思路**

### 代码风格

- C#：4 空格缩进，不写 nullable 禁用
- 注释优先解释"**为什么**"（尤其是 Word 的坑点，如"这里不能用 ArgumentList，PS 5.1 跑在 .NET Fx 4.x 上没这属性"）
- 不怕注释长——我自己半年后回来修 bug 时最感激的就是当年写的"为什么这么做"注释

### 避免踩这些坑

1. **bat 保持纯 ASCII + CRLF**。中文 Windows cmd 默认 GBK 码页，bat 文件里有中文字符会整行乱码解析成命令。
2. **ps1 必须 UTF-8 with BOM**。PowerShell 5.1 没 BOM 会按系统码页读文件，中文变乱码。
3. **不要用 `ProcessStartInfo.ArgumentList`**。PS 5.1 跑在 .NET Framework 4.x 上没这个属性。用 `& $exe @args` 调用运算符。
4. **PowerShell `.GetNewClosure()` 会吞 `$script:` 赋值**。`$script:var = $r` 在闭包里只改闭包副本。要共享状态，用一个外层的 hashtable（`$state = @{}`；闭包里 `$state.key = ...`）。
5. **rPr 里不要有 `*Theme` 属性**。Word 的主题字体优先级高于显式字体，留着会覆盖你设的字体。重建 rPr 时按 schema 顺序从零构造最干净。
6. **OpenXML schema 顺序**：`pStyle` 必须是 pPr 的第一个子元素；`rFonts` 通常应在 rPr 最前。
7. **`Indentation.FirstLineChars` 是 `Int32Value` 不是 string**（OpenXML SDK 3.x）。

### 加新功能指引

- **改字体/字号/行距默认值** → `src/MiniMaxAIDocx.Core/FormatOptions.cs` 的 `PresetDefault`
- **加新级标题（四级/五级）** → 详见 `docs/DEVELOPMENT.md` 的"加新样式"小节
- **改 TOC / 页码行为** → `GongWenFormatter.SetupSections`
- **加 CLI 参数** → `Program.cs` + `FormatOptions.cs` + (可选) `launcher/format.ps1`

### 测试

- 每个 fixture 测一个场景（中文编号 / 数字嵌套 / WPS 污染 / 自动编号多级 / 等）
- 加新代码路径时加对应 fixture
- `verify.py` 按角色逐字段断言（字体 / 字号 / 粗体 / 行距 / 缩进 / 对齐 / 无 theme 残留 / 无 numPr 残留 / 有 outlineLvl）

---

## English

Thanks for wanting to contribute! Notes below.

### Dev Environment

- **.NET 8 SDK** (required)
- **Python 3.8+** with `python-docx` and `lxml` (for tests)
- **PowerShell 5.1+** (Windows built-in) or **pwsh** 7.x (any OS)

On Linux without sudo:
```bash
curl -fsSL https://dot.net/v1/dotnet-install.sh | bash /dev/stdin --channel 8.0 --install-dir $HOME/.dotnet
export PATH="$HOME/.dotnet:$PATH"
```

### Layout

```
src/                            C# sources
launcher/                       Windows launchers (drag-drop + dialogs)
test/                           Test fixtures & runner
docs/                           Design docs
.github/workflows/              CI
```

### Flow

1. Fork, clone, hack
2. Local build: `./build.sh win-x64`
3. Run tests: `cd test && bash run_all_tests.sh` (all 7 fixtures must pass)
4. Submit PR with symptom → root cause → fix

### Pitfalls to Avoid

1. **.bat must be pure ASCII + CRLF.** Chinese Windows cmd defaults to GBK; non-ASCII bytes in a bat become mojibake parsed as commands.
2. **.ps1 must be UTF-8 *with* BOM.** Windows PowerShell 5.1 falls back to the system codepage without BOM.
3. **Don't use `ProcessStartInfo.ArgumentList`.** Missing in .NET Framework 4.x (where PS 5.1 runs). Use the call operator: `& $exe @args`.
4. **PowerShell's `.GetNewClosure()` swallows `$script:` writes.** Use a hashtable as shared container instead.
5. **Never emit `*Theme` attributes in rFonts.** Theme fonts override explicit fonts in Word.
6. **Mind the OpenXML schema order.** `pStyle` is pPr's first child; `rFonts` comes first inside rPr.
7. **`Indentation.FirstLineChars` is `Int32Value`**, not a string (OpenXML SDK 3.x).

### Style

- 4-space indentation
- Comment **why**, not what — especially for Word / WPS quirks
- Keep comments verbose; future-you will thank you

### Tests

Each fixture covers a scenario (Chinese-prefix / nested decimal / WPS pollution / auto multi-level numbering / etc.). Add a fixture when you add a new code path. `verify.py` asserts per-role: font / size / bold / line spacing / indent / alignment / no theme residue / no numPr residue / has outlineLvl.

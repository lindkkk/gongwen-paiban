# 公文排版 · gongwen-paiban

[English](#english) · [简体中文](#中文)

<p align="center">
  <strong>一键把 Word 文档排成符合中国公文 / 学术论文规范的格式</strong>
</p>

<p align="center">
  <a href="https://github.com/YOUR/REPO/actions"><img alt="Build" src="https://img.shields.io/badge/build-passing-brightgreen"></a>
  <a href="./LICENSE"><img alt="License" src="https://img.shields.io/badge/license-MIT-blue"></a>
  <a href="https://dotnet.microsoft.com/"><img alt=".NET 8" src="https://img.shields.io/badge/.NET-8.0-512BD4"></a>
  <img alt="Windows/Linux/macOS" src="https://img.shields.io/badge/platform-win--x64%20%7C%20linux--x64%20%7C%20osx--arm64-lightgrey">
</p>

---

## 中文

### 这是什么

一个**单文件、零依赖**的 Windows 命令行工具，把任意 `.docx` **按中国党政机关公文规范（GB/T 9704-2012 的常见简化版）或学术论文通行格式** 重新排版。自动识别：

- 文档主标题、一级/二级/三级标题、正文、脚注
- Word 内置 Heading 样式 / 段落大纲级别 / 自动编号 (`w:numPr`)
- 常见 WPS / Office 塞进去的主题字体、段落边框、`numPr` 列表缩进等"污染"

并把它们统一替换为规范格式。

支持：

- 字体字号 / 加粗斜体 / 对齐 / 行距（倍数或固定磅）/ 段前段后 / 首行缩进 的**每级独立配置**
- 用户**指定编号样例**（输入"一"、"1."、"(1)" 等即可，容错规则见下）
- 封面 / 目录 / 正文 分节页码（阿拉伯、罗马、无页码）
- `numPr` 自动编号"固化"成文字（删编号前先把 Word 要渲染的"1."、"一、" 等插进段首，不丢编号信息）

### 快速上手（Windows 普通用户）

1. 到 [Releases](https://github.com/YOUR/REPO/releases) 下最新 `gongwen-paiban-win-x64.zip`，解压
2. 把待排版的 `.docx` 文件**拖到 `format.bat`** 上
3. 依次回答三步对话框（文档来源 / 编号方式 / 样式），点"开始排版"
4. 同目录会生成 `<原文件名>_formatted.docx`

![drag-drop illustration](docs/screenshots/drag-drop.png) *（待补图）*

解压后三个文件必须放一起：

```
gongwen-paiban.exe    # 主程序（自包含，≈37MB）
format.bat            # 拖拽启动器
format.ps1            # 对话框逻辑
```

#### 功能截图 *（示意）*

| 主界面 | 样式编辑器 |
|---|---|
| `docs/screenshots/main.png` | `docs/screenshots/editor.png` |

### 命令行用法（Windows / Linux / macOS）

```bash
# 基本：自动识别 + 内置公文规范
gongwen-paiban format 输入.docx -o 输出.docx

# 指定来源和各级标题的编号样例（程序智能推导 regex）
gongwen-paiban format 输入.docx -o 输出.docx \
    --source wps \
    --h1-marker "一、"    \
    --h2-marker "（一）"  \
    --h3-marker "1."

# 自定义各级字体字号（JSON 配置）
gongwen-paiban format 输入.docx -o 输出.docx --config style.json

# 导出默认配置作模板
gongwen-paiban dump-config > my-style.json
# 编辑 my-style.json 后：
gongwen-paiban format 输入.docx -o 输出.docx --config my-style.json

# 测试编号样例的 regex 推导
gongwen-paiban test-marker "1.1" "1.1 节标题" "1.1.1 小节"

# 预览段落结构
gongwen-paiban preview 输入.docx
```

### 内置规范（可被 `--config` 覆盖）

| 元素 | 字体 | 字号 | 加粗 | 行距 | 首行缩进 | 对齐 |
|---|---|---|---|---|---|---|
| 主标题 | 方正小标宋简体 | 二号 22pt | 否 | 固定 28 磅 | — | 居中 |
| 一级 | 黑体 | 三号 16pt | 否 | 1.5 倍 | 2 字符 | — |
| 二级 | 楷体_GB2312 | 三号 16pt | 是 | 1.5 倍 | 2 字符 | — |
| 三级 | 仿宋_GB2312 | 三号 16pt | 是 | 1.5 倍 | 2 字符 | — |
| 正文 | 仿宋_GB2312 | 三号 16pt | 否 | 1.5 倍 | 2 字符 | — |
| 脚注 | 仿宋_GB2312 | 四号 14pt | 否 | 单倍 | 0 | — |
| 页码 | 仿宋_GB2312 | 小四 12pt | — | — | — | 居中 |

**页码**：封面无；目录 <2 页无；目录 ≥2 页罗马 I/II/III；正文从第 1 页起阿拉伯。

一级标题额外：段前段后 6 磅。

### 标题自动识别逻辑

按优先级：

1. **Word 既有样式**：`Heading 1/2/3`、`标题 1/2/3`、或 `w:outlineLvl=0/1/2`
2. **用户指定的编号样例**（`--h*-marker`）对"可见前缀 + 段落文本"做 regex 匹配
3. **自动编号的 `w:numPr`**：按 `ilvl=0/1/2` 映射 H1/H2/H3（仅对看起来像标题的短粗体段落）
4. **编号方案正则**：先扫全文判中文派（`一、/（一）/1.`）还是十进制派（`1./1.1/1.1.1`）再分级
5. **信号兜底**：短段 + 加粗 + 无句末标点 → H2

详见 [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)。

#### 编号样例的智能推理

你输入的简短样例会被自动扩展成宽松 regex：

| 你输入 | 匹配 |
|---|---|
| `一` | `一、` / `一.` |
| `1` / `1.` | `1.`（**不会错吃** `1.1`） |
| `1.1` | 恰 2 级（**不会错吃** `1.1.1`） |
| `1.1.1` | 恰 3 级 |
| `(1)` / `（1）` | 中英文括号任意组合 |
| `(一)` / `（一）` | 同上，汉字数字 |
| `第一章` / `第1章` | `第[一二…十]+章` / `第\d+章` |
| 乱码 `abc!@#` | 推不出 → 回退到自动识别（不崩） |

### 从源码构建

```bash
# 需要 .NET 8 SDK（Windows / Linux / macOS 均可）
git clone https://github.com/YOUR/REPO.git
cd gongwen-paiban

./build.sh win-x64      # 输出 dist/win-x64/gongwen-paiban.exe (+ 启动器)
# 或
./build.sh linux-x64
./build.sh osx-arm64
# Windows PowerShell:
.\build.ps1 win-x64
```

详见 [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md)。

### 测试

```bash
pip install python-docx lxml   # 测试依赖
cd test
bash run_all_tests.sh          # 7 个场景 fixture + 逐字段断言
```

### 已知限制

- 只处理 `.docx`。`.doc` 请先用 Word / WPS 另存为 `.docx`。
- 表格单元格内段落不格式化。
- Word 自动生成 TOC（`w:sdtBlock`）不重排。
- 字体效果：输出 XML 里字体名是正确的，但如果打开机器没装"方正小标宋简体"等字体，视觉会回退——换到装了字体的机器就恢复。

### 许可证

MIT License — 自由使用 / 修改 / 分发，见 [LICENSE](LICENSE)。

### 贡献

欢迎 issue 和 PR。贡献前请读 [CONTRIBUTING.md](CONTRIBUTING.md)。

---

## English

### What is this

A **zero-dependency single-file** command-line tool that reformats any `.docx` to match **Chinese official document conventions** (GB/T 9704-2012 commonly-adopted subset) or typical academic paper formatting. Automatically detects:

- Document title, H1/H2/H3 headings, body text, footnotes
- Word built-in Heading styles, paragraph outline levels, auto-numbering (`w:numPr`)
- Pollution from WPS / Office like theme fonts, paragraph borders, list indents

…and rewrites them in conformance. Supports per-level configuration of font/size/bold/italic/alignment/line-spacing/indent, user-specified numbering markers with fault-tolerant regex inference, multi-section page numbering (arabic / roman / none), and inlining of Word's auto-generated numbers into literal text.

### Quick Start (Windows end user)

1. Grab the latest `gongwen-paiban-win-x64.zip` from [Releases](https://github.com/YOUR/REPO/releases), extract it
2. **Drag** your `.docx` onto **`format.bat`**
3. Three-step wizard: source tool → numbering → style → Go
4. A `<name>_formatted.docx` is produced next to the original

The three files must stay together:

```
gongwen-paiban.exe    # ~37MB self-contained CLI
format.bat            # drag-drop launcher
format.ps1            # WinForms wizard
```

### CLI Usage

```bash
# Basic: auto-detect + built-in defaults
gongwen-paiban format input.docx -o output.docx

# Specify source and heading markers (tool infers flexible regex)
gongwen-paiban format input.docx -o output.docx \
    --source wps \
    --h1-marker "一、"     \
    --h2-marker "（一）"   \
    --h3-marker "1."

# Per-level style customisation via JSON
gongwen-paiban format input.docx -o output.docx --config style.json

# Export default config as template
gongwen-paiban dump-config > my-style.json

# Sanity-check a marker sample
gongwen-paiban test-marker "1.1" "1.1 Section" "1.1.1 Sub"
```

### Built-in Defaults (overridable via `--config`)

| Role | Font | Size | Bold | Line spacing | First-line indent | Align |
|---|---|---|---|---|---|---|
| Title | FZXiaoBiaoSong (方正小标宋简体) | 22pt | No | Exactly 28pt | — | Center |
| H1 | SimHei (黑体) | 16pt | No | 1.5× | 2 chars | — |
| H2 | KaiTi_GB2312 (楷体) | 16pt | Yes | 1.5× | 2 chars | — |
| H3 | FangSong_GB2312 (仿宋) | 16pt | Yes | 1.5× | 2 chars | — |
| Body | FangSong_GB2312 | 16pt | No | 1.5× | 2 chars | — |
| Footnote | FangSong_GB2312 | 14pt | No | Single | 0 | — |
| Page # | FangSong_GB2312 | 12pt | — | — | — | Center |

Page numbers: none on cover; none on TOC <2 pages; roman I/II/III on TOC ≥2 pages; arabic from body p.1.

### Heading Classification (priority order)

1. **Existing Word style** (`Heading 1/2/3`, `标题 1/2/3`, or `w:outlineLvl`)
2. **User-specified markers** (`--h*-marker`) applied to *visible prefix + text*
3. **Auto-numbering `w:numPr`** mapping `ilvl=0/1/2 → H1/H2/H3` for short bold paragraphs
4. **Numbering-scheme regex** (Chinese-style `一、/（一）/1.` vs. decimal `1./1.1/1.1.1`)
5. **Short + bold + no terminal punctuation → H2** (fallback)

See [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) for details.

### Build from Source

```bash
# Needs .NET 8 SDK
git clone https://github.com/YOUR/REPO.git
cd gongwen-paiban

./build.sh win-x64       # → dist/win-x64/gongwen-paiban.exe (+ launchers)
./build.sh linux-x64
./build.sh osx-arm64
# Or on Windows PowerShell:
.\build.ps1 win-x64
```

### Run Tests

```bash
pip install python-docx lxml
cd test && bash run_all_tests.sh
```

### Known Limitations

- `.docx` only. Convert `.doc` → `.docx` in Word/WPS first.
- Paragraphs inside table cells are left as-is.
- Word's auto-TOC (`w:sdtBlock`) is not re-styled.
- Font rendering falls back visually if the target machine lacks the referenced Chinese fonts, but the XML font name is correct — copy to a system with the fonts and it will render right.

### License

MIT — see [LICENSE](LICENSE).

### Contributing

Issues and PRs welcome. Please read [CONTRIBUTING.md](CONTRIBUTING.md).

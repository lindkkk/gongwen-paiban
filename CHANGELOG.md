# 变更日志

按时间顺序记录从起点到当前状态所有关键修复与功能演进。

---

## v2.2 — 2026-04-17 晚

### 修复

- **`ProcessStartInfo.ArgumentList` 在 Windows PowerShell 5.1 上不存在** `ps1`
  - 该属性是 .NET Core 2.1+ / .NET 5+ 才有；Win11 自带 PS 5.1 跑在 .NET Framework 4.x
  - 改成 `& $exe @exeArgs`（PS 调用运算符 + splat），5.1 / 7 都兼容
  - 现象：对话框走完点确定后 ps1 抛异常，退出码 99，日志里连 `[ps1 …]` 条都没有

- **bat 的 `1>>"%LOG%"` 长时间占用日志文件写锁** `bat`
  - ps1 在 PS 执行期间 `Add-Content` 同一文件时被 OS 拒绝，try/catch 吞掉后所有日志静默丢失
  - 改为 bat 不再对 powershell 重定向；日志文件完全交给 ps1 自己 `Add-Content`

- **bat 纯 ASCII 化 + CRLF 换行 + ps1 加 UTF-8 BOM** `bat` `ps1`
  - 现象：用户拖拽后 bat 窗口里一堆 `'xxx' 不是内部或外部命令` + 乱码
  - 根因：bat 原本是 UTF-8 无 BOM，中文 Windows cmd 默认 GBK 码页去解释 UTF-8 字节，乱码里含 `|` `>` `&` 等特殊符号被解析成管道
  - ps1 同理：PowerShell 5.1 没 BOM 默认按系统码页读 .ps1，中文字符串全变乱码
  - 修复：bat 内容全改英文（注释也英文），CRLF 行尾；ps1 写文件时前置 `EF BB BF`
  - bat 开头加 `chcp 65001 >nul` 统一输出编码

- **ps1 的 `-File` 参数名和 powershell.exe 的 `-File` 撞车** `ps1`
  - 原 `param([string]$File)` 被 `powershell -File script.ps1 -File "input.docx"` 的第二个 `-File` 吞掉
  - 改名为 `-InputDocx`，消除歧义

### 新增

- **交互式启动器** `ps1`
  - 拖 .docx 到 `format.bat` → PowerShell 弹三步对话框：
    - 第 1 步：文档来源（WPS / Office / 不确定）
    - 第 2 步：是否知道各级编号格式（是 / 否 / 取消）
    - 第 3 步：三格文本框让用户输入各级编号样例
  - 无论成败 bat 都 `pause` 等用户按键关窗
  - 每一步都写 `paiban-log.txt`，方便事后诊断

- **智能标题编号推理器** `MarkerPatternInferrer.cs`
  - 把用户随手输入的"标题编号样例"扩成宽松 regex
  - 容错：括号宽度错配 / 缺后缀标点 / 纯数字 vs 嵌套数字 自动区分
  - 关键：negative lookahead 锁死层级，`1.` 不会错吃 `1.1`，`1.1` 不会错吃 `1.1.1`
  - 无法解析的输入返回 null，回退到自动识别，不崩

- **CLI 新增参数** `Program.cs`
  - `--source {wps|office|auto}`（目前仅信息性，预留给将来做差异化处理）
  - `--h1-marker "一、"` / `--h2-marker "（一）"` / `--h3-marker "1."`

---

## v2.1 — 2026-04-17 下午

### 修复

- **自动编号自定义级别不再全被降为 H2** `GongWenFormatter.cs`
  - 新增 `DetectRoleFromNumPr`：段落有 `numPr` 且短+粗体时，按 `ilvl=0/1/2` 分别判 H1/H2/H3
  - 原先只要是短粗体就统一降 H2，导致用户用自动编号多级时三级全挤到二级

- **括号宽度 / 开闭错配兼容**
  - regex 字符类 `[（(]…[）)]` 独立处理开闭括号，`（1)` `(1）` 等混搭都认

### 新增

- **自动编号前缀固化成文本** `NumberingResolver.cs`
  - 段落若挂 `numPr`，按文档顺序跑一遍计数器，把 Word 本会动态渲染的 "1." / "一、" / "1.1" 等**字面化插入段落开头**
  - 因此删除 `numPr`（清除列表缩进污染）后不丢可见编号
  - 支持 numFmt: `decimal` / `decimalZero` / `upperLetter` / `lowerLetter` / `upperRoman` / `lowerRoman` / `chineseCounting` / `ideographDigital` 等
  - 多级嵌套（`%1.%2`/`%1.%2.%3`）正确解析

---

## v2.0 — 2026-04-17 中午

### 根本性修复：rPr/pPr 完全重建

用户报告"正文变宋体、二级标题缩进过度"。诊断后发现：

- **主题字体覆盖显式字体**：原稿 `rFonts` 里同时有 `ascii="FangSong_GB2312"` 和 `asciiTheme="minorEastAsia"`，Word 优先用后者，主题里 eastAsia 为空 → 回落到系统默认宋体
- **`numPr` 残留**：二级标题挂着列表编号，Word 按列表定义加额外缩进
- **pPr / rPr 里一堆 WPS 塞的脏数据**：`keepNext` / 边框 / `kinsoku` / `snapToGrid` / 段落标记 rPr / `kern` / `bdr` / `lang` 等
- **schema 违规**：`pStyle` 被 `AppendChild` 放到 pPr 末尾，schema 要求第一位
- **字体名用 Latin 写法**：`FangSong_GB2312` 在中文 Windows 上未必匹配以中文名注册的字体

**修复思路**：`RebuildPPr` 和 `RebuildRPr` 把原 pPr / rPr **整个删掉重建**，按 schema 顺序写只含必要子项的干净结构。字体名全换中文（`仿宋_GB2312` / `楷体_GB2312` / `黑体` / `方正小标宋简体`），加 `w:hint="eastAsia"`，**绝不写任何 `*Theme` 属性**。

---

## v1.2 — 2026-04-17 上午 · 页码修正

- 之前用 `titlePg=true` 隐藏正文首页页码，与用户"页码从正文首页开始标注"的规范不符
- 改为每个 section 明写 `<w:pgNumType w:start="1" fmt="..."/>`，**不再使用 `titlePg`**
- 正文首页直接显示 "1"

---

## v1.1 — 2026-04-17 上午 · TOC 边界检测

- 原目录项含 "一、" 等前缀被误当标题
- 改用点引线 `\.{3,}` / 省略号 / 末尾页码数字 作为 TOC 项特征，不再按标题前缀判

---

## v1.0 — 2026-04-17 上午 · 核心重写

起点：一份 skill 里的 `GongWenFormatter.cs` 只做了很粗糙的排版，用户担心格式规范度。诊断发现：

- 主标题 `GongWenTitle` 样式定义了但**从未被应用**
- 分类器只认 `一、` `（一）` `1.` 的**第一个**，后续同级全被漏
- 页码逻辑自相矛盾
- `body.MainDocumentPart` 在 OpenXML SDK 里根本不存在（编译错误）

**重写**：

- 新增 `DetectRoleFromExistingStyle`（读 Heading 1/标题1/outlineLvl）
- 新增编号方案检测（中文派 vs 十进制派）
- 重写分类器：**优先级 4 级**：既有样式 → numPr 层级 → 用户指定 marker → 编号方案正则 → 短粗体兜底
- Section 改造：封面 → TOC(Roman if ≥2 pages) → 正文，各自 `pgNumType`
- 加入 `ApplyTitleStyle` 并在识别到主标题时调用

---

## v0 — 起点

初始 skill 代码，存在多处 bug（见 v1.0 / v2.0 描述的"修复前"部分）。

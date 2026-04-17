# Changelog · 变更日志

All notable changes to this project are documented here.

---

## [v2.6.0] — 2026-04-17 · 开源 release 准备

仓库重组为开源项目常见布局，新增双语 README / CONTRIBUTING / GitHub Actions。

**目录重组**：

- `src/` — C# 源码（从 `.claude/skills/gongwen-paiban/scripts/dotnet/` 迁入）
- `launcher/` — Windows 启动器（从 `dist/win-x64/` 迁入，现在 `dist/` 全部是构建产物）
- `docs/` — 架构 / 开发 / 规范文档（ARCHITECTURE / DEVELOPMENT / format-spec）
- `test/` — 原位保留
- `build.sh` / `build.ps1` — 一键构建单文件自包含 exe
- `.github/workflows/build.yml` — push 时自动构建三平台（win/linux/mac），tag 时自动发 Release

**文档**：

- 根目录 **README.md**（中英双语门面）
- **CONTRIBUTING.md**（中英双语 + Word / WPS / PS 一系列踩坑指南）
- 更新 docs/ARCHITECTURE.md / DEVELOPMENT.md 用新路径

**skill 包装壳**：`.claude/skills/gongwen-paiban/` 保留但变成薄壳，主要文件都 link 到仓库根。

无功能改动。7 个 fixture 全绿，`build.sh` 跑通，单文件 exe 仍然 36 MB 可独立运行。

---

## [v2.5.2] — 2026-04-17 · 用户 marker 优先级 > numPr ilvl

**Bug**：用户原稿 H2 段落用 `w:numPr ilvl=0` 做自动编号，可见的"1." 是 Word 渲染的，段落文本里并没有"1."。用户填 `--h2-marker "1"` 期望按"1."识别为 H2，但 classifier 先走 numPr ilvl 路径把 `ilvl=0` 映射为 H1。

**修**：

1. `NumberingResolver` 移到 `Analyze` 之前跑，先算出每段 Word 要渲染的前缀
2. 调整优先级：既有样式 → **用户 marker** → numPr ilvl → 编号方案 → 短粗体
3. 用户 marker 匹配时用 `prefix + " " + text` 作为文本对象，让"一、首章"这种只有 numPr 的段也能被"一"这种 marker 命中

## [v2.5.1] — 2026-04-17 · WPS 导航窗格仍错归的终极修复

v2.5 给 pPr / 样式都加了 `outlineLvl` 但 WPS 导航窗格依然把 H2 归到顶层。WPS 不光看 outlineLvl，还要求样式具备 Word 内置 heading 的"指纹"。

修：标题类直接复用 Word 标准 styleId（`Title` / `Heading1` / `Heading2` / `Heading3`），name 用 `heading 1/2/3`（小写带空格），补上 `<w:qFormat/>` 和 `uiPriority`。Chinese 显示名塞进 `<w:aliases>`。

## [v2.5] — 2026-04-17 · 标题大纲级别

给所有标题段和标题样式写 `<w:outlineLvl w:val="N"/>`：H1=0 / H2=1 / H3=2。Word 导航窗格按大纲级别分类。

## [v2.4.3] — 2026-04-17 · UI：去掉"我知道编号"checkbox

编号输入框常时可用，填了就用、空=自动。避免用户忘勾选 checkbox 导致 marker 失效。

## [v2.4.2] — 2026-04-17 · 两个 UI bug

1. 段后 NumericUpDown 被 Label 覆盖 20px，点不到 → 重排坐标
2. 样式编辑器确定后改动对输出无效 → PowerShell `.GetNewClosure()` 吞 `$script:` 赋值，改用共享 hashtable 作状态容器

## [v2.4.1] — 2026-04-17 · 模态对话框沉任务栏

主窗口 `TopMost=true` + ShowDialog 没传 Owner → 样式编辑器一闪沉任务栏。去掉 TopMost，传 Owner 建立父子模态。

## [v2.4] — 2026-04-17 · 单窗口主 UI

4 个连续对话框合为 1 个主窗口 + 3 个 GroupBox + 可选模态样式编辑器。

## [v2.3] — 2026-04-17 · 每角色自定义样式

新增 `StyleSpec` / `FormatOptions`，支持 JSON 配置、每角色独立字体 / 字号 / 加粗 / 斜体 / 对齐 / 行距（倍数 or 固定磅）/ 段前 / 段后 / 首行缩进。CLI 新增 `--config` / `dump-config`。默认行为向后兼容。

## [v2.2] — 2026-04-17 · 交互式 PowerShell 启动器

拖拽 `.docx` → `format.bat` → PowerShell 弹三步对话框（来源 / 编号方式 / 编号样例）。诸多踩坑：
- bat 纯 ASCII + CRLF（中文 Windows cmd GBK 码页不认 UTF-8 bat）
- ps1 UTF-8 BOM + CRLF（PS 5.1 无 BOM 按系统码页读）
- `ProcessStartInfo.ArgumentList` 在 .NET Framework 4.x 上不存在 → 改用 `& $exe @args`
- ps1 参数 `-File` 名和 powershell.exe 自身的 `-File` 撞车 → 改名 `-InputDocx`
- bat 对 powershell 重定向会占锁 → 不重定向让 ps1 写自己的日志

新增 `MarkerPatternInferrer`：用户随手输入的"编号样例"扩成宽松 regex（`"一"`→`^[一二…十]+[、.．]`、`"1.1"`→`^\d+[.]\d+(?![.]\d)` 锁定层级）。

## [v2.1] — 2026-04-17 · H1/H3 自动编号

分类器新增 `DetectRoleFromNumPr`：段落挂 numPr 且短粗体时按 ilvl 判 H1/H2/H3。括号宽度中英文混搭也认。

`NumberingResolver`：跑计数器把 Word 要渲染的"1." "一、" "（1）" 固化成段首文本，删 numPr 后不丢编号。

## [v2.0] — 2026-04-17 · rPr / pPr 完全重建

正文变宋体、二级标题缩进过度。根因：
- rFonts 带 `*Theme` 属性压过显式字体
- numPr 残留引入列表缩进
- pPr / rPr 里一堆 WPS 塞的脏数据

修：`RebuildPPr` / `RebuildRPr` 整个 pPr / rPr 节点删掉重建，按 schema 顺序写只含必要子项的干净结构。字体全换中文名（`仿宋_GB2312` 等），加 `w:hint="eastAsia"`，**绝不写 `*Theme`**。

## [v1.x] — 2026-04-17 上午

- v1.2 页码：不再用 `titlePg`，每 section 写 `pgNumType start=1`
- v1.1 TOC 边界：按点引线 / 省略号 / 末尾页码判，不按标题前缀
- v1.0 核心重写：多级标题分类器、编号方案检测、主标题应用、section/页码架构

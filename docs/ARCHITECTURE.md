# 架构与数据流

## 一句话定位

给一份 `.docx`，按预设的公文/论文排版规范，**识别出它的主标题、一到三级标题、正文、脚注**，并重写字体字号行距段落属性与页码，同时清掉 WPS 等编辑器遗留的样式污染。

## 运行时拓扑

```
┌─────────────────┐   拖拽.docx      ┌─────────────────┐
│  format.bat     ├─────────────────▶│  format.ps1     │
│  (ASCII+CRLF)   │                  │  (UTF-8 BOM)    │
└─────────────────┘                  │  WinForms 对话框 │
                                     └────────┬────────┘
                                              │ CLI args
                                              ▼
                                     ┌────────────────────┐
                                     │ gongwen-paiban.exe │
                                     │  (self-contained   │
                                     │   net8.0 win-x64)  │
                                     └────────┬───────────┘
                                              │
                                     reads & writes .docx
                                              │
                                              ▼
                                     ┌────────────────────┐
                                     │  DocumentFormat    │
                                     │  .OpenXml 3.2      │
                                     └────────────────────┘
```

一切在本地完成，无网络依赖。

## 源码目录

```
src/                                         # C# 源码
├── MiniMaxAIDocx.slnx
├── MiniMaxAIDocx.Core/
│   ├── GongWenFormatter.cs                  # ★ 主入口，分类器 / 样式应用 / section 设置
│   ├── NumberingResolver.cs                 # numPr 自动编号前缀字面化
│   ├── MarkerPatternInferrer.cs             # 用户标题样例 → regex
│   ├── FormatOptions.cs / StyleSpec.cs      # CLI 选项 + 单角色规格
│   └── MiniMaxAIDocx.Core.csproj
└── MiniMaxAIDocx.Cli/
    ├── Program.cs                           # 命令行解析
    └── MiniMaxAIDocx.Cli.csproj
launcher/
├── format.bat                               # Windows 拖拽启动器（ASCII + CRLF）
└── format.ps1                               # WinForms 对话框（UTF-8 BOM + CRLF）
docs/
├── ARCHITECTURE.md                          # 本文档
├── DEVELOPMENT.md                           # 开发 / 构建 / 测试
└── format-spec.md                           # 规范文字版，和代码常量严格对齐
test/
├── gen_*.py                                 # 各种场景的测试文档生成器
├── verify.py                                # 严格校验输出文档
└── run_all_tests.sh
dist/                                        # 构建产物（.gitignore）
build.sh / build.ps1                         # 一键构建脚本
.github/workflows/build.yml                  # CI：push/tag 自动构建 + 发 release
.claude/skills/gongwen-paiban/               # Claude Code skill 包装壳（薄）
```

## 核心数据流：一次 format 调用

```
WordprocessingDocument.Open
        │
        ▼
CreateGongWenStyles(mainPart)         ←  注入 6 个自定义 styleId 到 styles.xml
        │
        ▼
Analyze(topLevelParas)                ←  结构分析 + 分类
        │
        ▼
NumberingResolver.Load + AdvanceAndFormat
  （按文档顺序跑计数器，给每段记下要固化的编号前缀）
        │
        ▼
ApplyFormatting(structure)            ←  RebuildPPr + RebuildRPr + InlineNumberPrefix
        │
        ▼
ApplyFootnoteFormatting(mainPart)     ←  脚注用同一 rebuild 路径
        │
        ▼
SetupSections(mainPart, body, s)      ←  封面 / TOC / 正文三类 section + 页脚
        │
        ▼
mainPart.Document.Save()
```

## 段落分类器（ClassifyParagraph）

**按优先级**逐级下落。命中即停。

| 优先级 | 判据 | 由谁实现 |
|---|---|---|
| 1a | Word 既有样式：`pStyle = Heading 1/2/3` 或 `标题1/2/3`，或 `outlineLvl=0/1/2` | `DetectRoleFromExistingStyle` |
| 1b | 段落挂 `numPr` 且短(≤50 字)+加粗，按 `ilvl=0/1/2` 映射 H1/H2/H3 | `DetectRoleFromNumPr` |
| 1c | 用户提供的 marker 样例生成的 regex（H3 → H2 → H1 顺序试） | `MarkerPatternInferrer` + 主循环 |
| 2  | 目录区 (`目录` / `Contents`) 和目录项（点引线/省略号/末尾页码） | `DetectToc` / `IsTocItemLike` |
| 3  | 主标题：TOC 前第一个非空短段（<60 字且不像标题前缀） | `DetectTitle` |
| 4  | 编号方案自检：全文扫"中文派 vs 十进制派"，选 regex 组 | `DetectNumberScheme` + `ClassifyHeadingByText` |
| 5  | 短+粗体+无句末标点兜底 → H2 | `IsShortBoldHeading` |
| 6  | 否则 Body | 默认 |

Cover / TOC 项**不应用我方样式**（保持原稿）。Title / H1 / H2 / H3 / Body / Footnote **都走 RebuildPPr + RebuildRPr 重建**。

## 为什么要"整个重建 pPr / rPr"

原本只删几个已知子元素不够。WPS / 旧 Office 会塞大量私有属性，尤其是：

1. **主题字体属性** `w:asciiTheme="minorEastAsia"` 等，**优先级高于显式字体**
2. **`w:numPr`** 自动列表，带自己的缩进定义
3. **段落标记 rPr**（pPr 内嵌的 rPr），影响段落标记和某些继承行为
4. 各类无害但花哨的属性：`keepNext` / `pBdr` / `kinsoku` / `snapToGrid` / `adjustRightInd` / `textAlignment` …

单独删不胜删。**直接把 pPr / rPr 整个节点扔掉、按 schema 顺序重新 new 一份只含必要子项的**，最稳。

```csharp
para.ParagraphProperties?.Remove();
var pPr = new ParagraphProperties();
pPr.AppendChild(new ParagraphStyleId { Val = styleId });   // schema 要求 pStyle 第一位
pPr.AppendChild(new Justification { ... });
pPr.AppendChild(new SpacingBetweenLines { ... });
pPr.AppendChild(new Indentation { FirstLineChars = 200 });
para.InsertAt(pPr, 0);
```

rPr 同理。关键是 rFonts 写中文字体名 + `w:hint="eastAsia"` + **绝不写任何 `*Theme` 属性**。

## NumberingResolver：自动编号前缀文本化

Word 的 `<w:numPr numId=X ilvl=Y/>` 不存储可见编号文本，Word 渲染时按 `numbering.xml` 的 `abstractNum` + `lvlText` + `numFmt` 动态算出来。

我们删 `numPr`（消除列表缩进污染）前先：

1. 读 `numbering.xml`，建 `abstractNumId → [LvlDef{lvlText, numFmt, start}]` 映射和 `numId → abstractNumId` 映射
2. 按文档顺序遍历所有段落，对每个有 `numPr` 的段：
   - 推进该 numId 下 ilvl 的计数器；比 ilvl 深的全部归零（Word 行为）
   - 按 lvlText (如 `"%1."` / `"%1、"` / `"（%3）"`) 替换 `%N` 为第 N 层当前计数器，按对应 numFmt 格式化
3. 把算出的字符串（如 `"1."`、`"一、"`、`"（3）"`）挂到 `ClassifiedPara.NumberPrefix`

应用样式后，`InlineNumberPrefix` 把这串文本前置插入段落开头（同一个刚重建的干净 rPr 下，自然继承标题字体）。

## MarkerPatternInferrer：用户样例智能 regex

原则：用户随手输，程序容错补齐。

```
"一"         →  ^[一二三四五六七八九十百千两]+[、．.]
"一、"       →  ^[一二三四五六七八九十百千两]+、
"1"          →  ^\d+[．.、](?!\d)     （负向先行断言阻止吃 "1.1"）
"1."         →  ^\d+[．.](?!\d)
"1.1"        →  ^\d+[．.]\d+(?![．.]\d)
"1.1.1"      →  ^\d+[．.]\d+[．.]\d+(?![．.]\d)
"(1)"/"（1）" →  ^[（(]\d+[）)]
"第一章"     →  ^第[一二三四五六七八九十百千两]+章
"abc!@#"     →  null → 回退自动识别
```

推理失败返回 null，不抛异常。

## Section / 页码

三类 section，按文档层次自低到高：

| section | 条件 | footerReference | pgNumType |
|---|---|---|---|
| 封面 | 检测到 TOC 时，TOC 之前的段 | 空页脚 | `start=1` 无 format |
| TOC | TOC 估算 ≥2 页时 | Roman 页脚 | `start=1 format=upperRoman` |
| 正文 | 总是 | 阿拉伯页脚 | `start=1 format=decimal` |

TOC <2 页时并入封面（两者都无页码）。封面独立 FooterPart，保证空内容不被继承。

**不使用 `w:titlePg`**——早期版本用它隐藏首页页码，与规范"页码从正文首页开始标注"不符。

## 字号 / 行距 / 缩进对照

常量集中在 `GongWenFormatter.cs` 顶部。改值只要改一行。

| 字号 | 磅值 | w:sz |
|---|---|---|
| 二号 | 22 | 44 |
| 三号 | 16 | 32 |
| 四号 | 14 | 28 |
| 小四 | 12 | 24 |

| 行距类型 | w:line | w:lineRule |
|---|---|---|
| 1.5 倍 | 360 | auto |
| 单倍 | 240 | auto |
| 固定 28 磅（用于主标题） | 560 | exact |

| 缩进 | w:firstLineChars |
|---|---|
| 2 字符 | 200 |
| 0 字符 | 0 |

段前段后 6 磅 = 120 DXA。

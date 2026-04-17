---
name: gongwen-paiban
license: MIT
metadata:
  version: "2.0.0"
  category: document-processing
  author: MiniMaxAI
description: >
  公文 / 学术论文排版工具。按预设规范对 Word 文档进行标准化排版：
  自动识别封面、目录、多级标题，重写字体字号行距首行缩进，并按规则设置页码。
triggers:
  - 公文排版
  - 排版
  - 格式排版
  - 政府公文
  - 行政公文
  - 红头文件
  - 论文排版
  - doc排版
---

# 公文排版工具 (gongwen-paiban)

对 Word 文档（.docx）进行一键标准化排版。

## 一、排版规范（硬性，工具严格执行）

### 1.1 字体字号 / 段落

| 角色 | 字体 | 字号 | 加粗 | 行距 | 首行缩进 | 其他 |
|------|------|------|------|------|----------|------|
| 标题 | 方正小标宋简体 `FZXiaoBiaoSong-B05S` | 二号 22pt | 否 | 固定值 28 磅 | — | 居中 |
| 一级标题 | 黑体 `SimHei` | 三号 16pt | 否 | 1.5 倍 | 2 字符 | 段前段后 6 磅 |
| 二级标题 | 楷体 `KaiTi_GB2312` | 三号 16pt | 是 | 1.5 倍 | 2 字符 | — |
| 三级标题 | 仿宋 `FangSong_GB2312` | 三号 16pt | 是 | 1.5 倍 | 2 字符 | — |
| 正文 | 仿宋 `FangSong_GB2312` | 三号 16pt | 否 | 1.5 倍 | 2 字符 | — |
| 脚注 | 仿宋 `FangSong_GB2312` | 四号 14pt | 否 | 单倍 | 0 | — |
| 页码 | 仿宋 `FangSong_GB2312` | 小四 12pt | — | — | — | 居中 |

### 1.2 页码规则

| 位置 | 规则 |
|------|------|
| 封面 | 无页码 |
| 目录（<2 页） | 无页码（并入封面 section） |
| 目录（≥2 页） | 罗马数字 I、II、III…，首页从 I 开始 |
| 正文 | 阿拉伯数字，**首页即显示 "1"** |
| 封底 / 空白页 | 无页码（交由用户通过分节保持） |

## 二、自动结构识别逻辑

工具按以下优先级识别段落角色：

1. **已有 Word 样式** — 若段落引用 `Heading 1/标题1`、`Heading 2/标题2`、`Heading 3/标题3`（或 `w:outlineLvl 0/1/2`），直接按层级套用。
2. **目录区** — 检测"目录 / 目 录 / Contents / Table of Contents"段作为 TOC 起点；其后连续"带点引线、省略号或末尾页码"的段落判为目录项，不重排样式。
3. **主标题** — TOC 之前的首个非空短段（<60 字且不像标题前缀）判为主标题。
4. **编号方案** — 扫描全文判断是中文编号（`一、/（一）/1.`）还是十进制编号（`1./1.1/1.1.1`），以此决定后续正则分级。
5. **三级标题正则**（中文方案）：
   - 一级：`^[一二三四五六七八九十百]+[、．.]` / `^第\d+[章篇部]` / `^#\s`
   - 二级：`^[（(][一二三四五六七八九十]+[）)]` / `^第\d+节` / `^##\s`
   - 三级：`^\d+[．.、]\s*[^\d\s]` / `^[（(]\d+[）)]` / `^第\d+条` / `^###\s`
6. **信号兜底** — 段落 ≤30 字、带粗体、无句末标点 → 识别为二级标题。典型用例：没有编号的"参考文献 / 附录 / 结语"。

### 识别局限

- TOC 若被 `w:sdtBlock`（Word 自动生成目录）包裹，当前不会识别内部段落；本工具不会破坏它但也不会对其重排。
- 表格单元格内的段落**不会**被重排（保持原样）。
- 脚注需是 Word 自带脚注（`footnotes.xml`）才会被格式化；段落末尾手打的注释不会识别。

## 三、使用方法

### 3.1 Windows 普通用户（推荐）

1. 下载 `dist/win-x64/` 目录
2. 把待排版的 `.docx` 文件**直接拖到** `排版此文件.bat` 上
3. 同目录下会生成 `<原文件名>_已排版.docx`

无需安装 .NET、Word、Python 或任何依赖。Windows 7 及以上双击即用。

### 3.2 命令行（Windows / Linux / macOS）

```bash
gongwen-paiban format 输入文件.docx -o 输出文件.docx
gongwen-paiban preview 输入文件.docx        # 列出段落文本
```

自行从源码构建：

```bash
cd scripts/dotnet
dotnet publish MiniMaxAIDocx.Cli/MiniMaxAIDocx.Cli.csproj \
    -c Release -r win-x64 --self-contained true \
    -p:PublishSingleFile=true -p:EnableCompressionInSingleFile=true \
    -o ../../../dist/win-x64
```

### 3.3 .doc 文件

本工具只处理 `.docx`。`.doc` 请先用 Word / WPS 打开 → 另存为 `.docx`。

### 3.4 字体

排版输出的文档里写的是"方正小标宋 / 仿宋_GB2312 / 楷体_GB2312 / 黑体"这些标准公文字体。Windows 默认自带 黑体 与 仿宋/楷体（标准版可能叫 "仿宋" 而不是 "仿宋_GB2312"——在 Word 里会自动回退）。方正小标宋简体需单独安装，未安装时 Word 会回退到相近字体，但文档 XML 写的是正确字体名，换到装了字体的机器上即刻正确。

## 四、验证

跑 `test/run_all_tests.sh` 会用三份自动生成的样本（中文编号+短 TOC、中文编号+长 TOC、无结构纯正文）跑一遍格式化 → 用 `test/verify.py` 对每个段落逐条校验字体 / 字号 / 粗体 / 行距 / 首行缩进 / section 页码类型。

## 五、目录结构

```
gongwen-paiban/
├── SKILL.md                    # 本文档
├── README.md                   # 简介
├── LICENSE
├── references/
│   └── gongwen_format_rules.md # 规范技术细节
└── scripts/
    ├── setup.sh / setup.ps1
    ├── env_check.sh
    ├── docx_preview.sh
    └── dotnet/
        ├── MiniMaxAIDocx.slnx
        ├── MiniMaxAIDocx.Core/
        │   └── GongWenFormatter.cs   # 核心逻辑
        └── MiniMaxAIDocx.Cli/
            └── Program.cs
```

## 六、许可证

MIT

# 公文排版格式规范（技术细节）

本文档列出各项排版参数与 OpenXML 字段值的对应关系，与 `GongWenFormatter.cs` 保持一致。

## 一、字体字号

### 1.1 标题

| 属性 | 要求 | OpenXML |
|------|------|----------|
| 字体 | 方正小标宋简体 | `<w:rFonts w:eastAsia="FZXiaoBiaoSong-B05S" w:ascii="FZXiaoBiaoSong-B05S" w:hAnsi="FZXiaoBiaoSong-B05S" w:cs="FZXiaoBiaoSong-B05S"/>` |
| 字号 | 二号 22pt | `<w:sz w:val="44"/>` |
| 粗体 | 否（显式 false） | `<w:b w:val="false"/>` |
| 对齐 | 居中 | `<w:jc w:val="center"/>` |
| 行距 | 固定值 28 磅 | `<w:spacing w:line="560" w:lineRule="exact"/>` |

### 1.2 一级标题

| 属性 | 要求 | OpenXML |
|------|------|----------|
| 字体 | 黑体 | `SimHei` |
| 字号 | 三号 16pt | `<w:sz w:val="32"/>` |
| 粗体 | 否 | `<w:b w:val="false"/>` |
| 段前段后 | 6 磅（120 DXA） | `<w:spacing w:before="120" w:after="120"/>` |
| 行距 | 1.5 倍 | `<w:spacing w:line="360" w:lineRule="auto"/>` |
| 首行缩进 | 2 字符 | `<w:ind w:firstLineChars="200"/>` |

### 1.3 二级标题

| 属性 | 要求 | OpenXML |
|------|------|----------|
| 字体 | 楷体_GB2312 | `KaiTi_GB2312` |
| 字号 | 三号 16pt | `<w:sz w:val="32"/>` |
| 粗体 | 是 | `<w:b/>` |
| 行距 | 1.5 倍 | `<w:spacing w:line="360" w:lineRule="auto"/>` |
| 首行缩进 | 2 字符 | `<w:ind w:firstLineChars="200"/>` |

### 1.4 三级标题

| 属性 | 要求 | OpenXML |
|------|------|----------|
| 字体 | 仿宋_GB2312 | `FangSong_GB2312` |
| 字号 | 三号 16pt | `<w:sz w:val="32"/>` |
| 粗体 | 是 | `<w:b/>` |
| 行距 | 1.5 倍 | 同上 |
| 首行缩进 | 2 字符 | 同上 |

### 1.5 正文

| 属性 | 要求 | OpenXML |
|------|------|----------|
| 字体 | 仿宋_GB2312 | `FangSong_GB2312` |
| 字号 | 三号 16pt | `<w:sz w:val="32"/>` |
| 粗体 | 否 | `<w:b w:val="false"/>` |
| 行距 | 1.5 倍 | `line=360 lineRule=auto` |
| 首行缩进 | 2 字符 | `firstLineChars=200` |

### 1.6 脚注

| 属性 | 要求 | OpenXML |
|------|------|----------|
| 字体 | 仿宋_GB2312 | `FangSong_GB2312` |
| 字号 | 四号 14pt | `<w:sz w:val="28"/>` |
| 行距 | 单倍 | `line=240 lineRule=auto` |
| 首行缩进 | 0 | `firstLineChars=0 firstLine=0` |

## 二、页码

### 2.1 规则

| 位置 | 规则 | 实现方式 |
|------|------|----------|
| 封面 | 无页码 | section 的 footerReference 指向 "空页脚" FooterPart |
| 目录 <2 页 | 无页码 | 并入封面 section |
| 目录 ≥2 页 | 罗马 I、II、III… | 独立 section + footerReference(Roman) + `<w:pgNumType w:start="1" w:fmt="upperRoman"/>` |
| 正文 | 阿拉伯数字 1、2、3…，**首页显示** | 独立 section + footerReference(Arabic) + `<w:pgNumType w:start="1" w:fmt="decimal"/>` |

### 2.2 字体字号

- 字体：仿宋_GB2312
- 字号：小四 12pt（`w:sz w:val="24"`）
- 对齐：居中（`jc w:val="center"`）

### 2.3 Footer 字段代码

```xml
<!-- 阿拉伯 -->
<w:fldChar w:fldCharType="begin"/>
<w:instrText> PAGE </w:instrText>
<w:fldChar w:fldCharType="separate"/>
<w:t>1</w:t>
<w:fldChar w:fldCharType="end"/>

<!-- 罗马 -->
<w:instrText> PAGE \* ROMAN </w:instrText>
```

### 2.4 关键：不使用 titlePg

历史上很多类似工具用 `<w:titlePg w:val="true"/>` 来隐藏正文首页页码。本工具**不使用 titlePg**，而是在每个 section 都写 `<w:pgNumType w:start="1"/>` 让页码从各 section 首页起重新计数——这样正文第 1 页就是 "1"，符合用户规范"页码从正文首页开始标注"。

## 三、自定义样式 ID

| StyleId | 中文名 | 用途 |
|---------|--------|------|
| `GongWenTitle` | 公文标题 | 主标题 |
| `Heading1GongWen` | 一级标题 | 一级 |
| `Heading2GongWen` | 二级标题 | 二级 |
| `Heading3GongWen` | 三级标题 | 三级 |
| `GongWenBody` | 公文正文 | 正文 |
| `GongWenFootnote` | 公文脚注 | 脚注 |

## 四、多 Section 架构

```
┌── 封面 Section ──────────────────────────────┐
│ footerReference → 空页脚                    │
│ pgNumType start=1（不显示，仅重置计数）     │
│ sectionType = nextPage                      │
└──────────────────────────────────────────────┘
         ↓（强制换页）
┌── 目录 Section（仅 ≥2 页才独立创建） ────────┐
│ footerReference → 罗马页脚                  │
│ pgNumType start=1 format=upperRoman         │
│ sectionType = nextPage                      │
└──────────────────────────────────────────────┘
         ↓
┌── 正文 Section（Body 末尾 sectPr） ──────────┐
│ footerReference → 阿拉伯页脚                │
│ pgNumType start=1 format=decimal            │
└──────────────────────────────────────────────┘
```

## 五、字号对照

| 字号 | 磅值 | `w:sz` |
|------|------|--------|
| 二号 | 22pt | 44 |
| 三号 | 16pt | 32 |
| 四号 | 14pt | 28 |
| 小四 | 12pt | 24 |

## 六、行距 / 间距对照

| 类型 | `w:line` | `w:lineRule` |
|------|----------|--------------|
| 1.5 倍 | 360 | auto |
| 单倍 | 240 | auto |
| 标题固定值 28 磅 | 560 | exact |

| 磅值 | DXA |
|------|-----|
| 6 磅 | 120 |

| 字符数 | `firstLineChars` |
|--------|-----------------|
| 0 字符 | 0 |
| 2 字符 | 200 |

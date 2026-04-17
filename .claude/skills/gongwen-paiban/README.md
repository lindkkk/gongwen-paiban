# 公文排版工具 (gongwen-paiban)

对 Word 文档（.docx）进行一键标准化排版。自动识别封面、目录、多级标题，重写字体字号行距首行缩进，并按规则设置页码。

## 快速开始（Windows 普通用户）

1. 取 `dist/win-x64/` 目录（含 `gongwen-paiban.exe` 和 `排版此文件.bat`）
2. 把 `.docx` 文件拖到 `排版此文件.bat` 上
3. 旁边会多出 `<原文件名>_已排版.docx`

零依赖，Windows 7 及以上可用。

## 命令行

```bash
gongwen-paiban format 输入.docx -o 输出.docx
gongwen-paiban preview 输入.docx
```

## 排版规则（摘要）

| 元素 | 字体 | 字号 | 备注 |
|------|------|------|------|
| 标题 | 方正小标宋简体 | 二号 | 居中、不加粗 |
| 一级 | 黑体 | 三号 | 段前段后6磅、不加粗 |
| 二级 | 楷体_GB2312 | 三号 | 加粗 |
| 三级 | 仿宋_GB2312 | 三号 | 加粗 |
| 正文 | 仿宋_GB2312 | 三号 | 1.5倍行距、首行缩进2字 |
| 脚注 | 仿宋_GB2312 | 四号 | 单倍、不缩进 |
| 页码 | 仿宋_GB2312 | 小四 | 居中 |

页码：封面无、目录<2页无、目录≥2页用罗马、正文首页起阿拉伯数字。

详细规则与自动识别机制见 [SKILL.md](SKILL.md)。

## 技术栈

- .NET 8 + DocumentFormat.OpenXml 3.2
- 单文件 self-contained 发布（Windows x64 ≈ 36MB）

## 许可证

MIT

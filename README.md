# docx-cn-skill — 中文 Word 文档生成工具

基于 python-docx 的专业中文文档生成工具，Claude Code docx skill 的中文增强模块。

## 功能

- 🀄 **中文排版规范** — 符合 GB/T 9704（党政公文）、GB/T 7713（学术论文）标准
- 📝 **字体层级体系** — 黑体/宋体/楷体/仿宋 + Times New Roman 混排
- 📄 **丰富模板库** — 党政公文、学术论文、调研报告、工作总结、日常文书
- 🎯 **零额外依赖** — 中文字体（宋体/黑体/楷体/仿宋）Windows 系统内置
- 🔧 **灵活定制** — 提供 ChineseDocument 工厂类和底层格式化函数

## 安装

```bash
pip install python-docx lxml
```

将此模块添加到 Claude Code docx skill 中：

```bash
# 复制中文模块到 skill 目录
cp scripts/chinese_format.py ~/.agents/skills/docx/scripts/
cp -r scripts/docx_gen/ ~/.agents/skills/docx/scripts/
```

## 快速开始

### 使用模板

```python
from scripts.docx_gen.official import generate_tongzhi

generate_tongzhi(
    output_path="通知.docx",
    title="关于做好端午节放假期间安全工作的通知",
    issuing_authority="XX大学学生工作处",
    doc_number="学工发〔2026〕15号",
    recipients="各学院：",
    body_paragraphs=["根据学校安排...", "一、加强安全教育...", "二、..."]
)
```

### 手动创建文档

```python
from scripts.docx_gen.base import ChineseDocument
from scripts.chinese_format import PAGE_DEFAULT

doc = ChineseDocument.create(PAGE_DEFAULT)
ChineseDocument.add_title(doc, "我的文档")
ChineseDocument.add_heading(doc, "第一章", level=1)
ChineseDocument.add_body(doc, "正文内容...")
ChineseDocument.save(doc, "output.docx")
```

## 模板目录

| 类别 | 模板 | 函数 | 参考标准 |
|------|------|------|---------|
| 党政公文 | 通知 | `generate_tongzhi()` | GB/T 9704 |
| 党政公文 | 请示 | `generate_qingshi()` | GB/T 9704 |
| 党政公文 | 报告 | `generate_baogao()` | GB/T 9704 |
| 党政公文 | 函 | `generate_han()` | GB/T 9704 |
| 学术论文 | 课程论文 | `generate_course_paper()` | GB/T 7713 |
| 学术论文 | 学位论文 | `generate_thesis()` | GB/T 7713 |
| 学术论文 | 开题报告 | `generate_proposal()` | GB/T 7713 |
| 日常文书 | 申请书 | `generate_application()` | 通用 |
| 日常文书 | 会议纪要 | `generate_meeting_minutes()` | 通用 |

## 中文排版规范

### 字体层级（黑体/宋体规则）

| 元素 | 中文字体 | 西文字体 | 字号 | 加粗 |
|------|---------|---------|------|------|
| 文档标题 | 黑体 | Times New Roman | 22pt (二号) | 是 |
| 一级标题 | 黑体 | Times New Roman | 16pt (三号) | 是 |
| 二级标题 | 黑体 | Times New Roman | 15pt (小三) | 是 |
| 三级标题 | 黑体 | Times New Roman | 12pt (小四) | 是 |
| 正文 | 宋体 | Times New Roman | 12pt (小四) | 否 |
| 题注 | 楷体 | Arial | 10.5pt (五号) | 否 |
| 公文正文 | 仿宋 | Times New Roman | 16pt (三号) | 否 |

### 页面设置

| 预设 | 上 | 下 | 左 | 右 | 适用 |
|------|---|---|---|---|---|
| PAGE_DEFAULT | 2.54cm | 2.54cm | 3.18cm | 3.18cm | 通用 |
| PAGE_OFFICIAL | 3.7cm | 3.5cm | 2.8cm | 2.6cm | 公文 |
| PAGE_ACADEMIC | 2.5cm | 2.5cm | 2.5cm | 2.5cm | 论文 |

## 依赖

- python-docx >= 1.1.0
- lxml >= 5.0.0
- 中文字体（Windows 内置）：宋体、黑体、楷体、仿宋

## 前提条件

此模块是 [Claude Code docx skill](https://claude.ai/code) 的中文增强模块。需要先安装基础 docx skill 才能使用编辑功能。仅文档生成功能可独立使用。

## 许可

MIT License

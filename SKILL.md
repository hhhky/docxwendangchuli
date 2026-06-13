---
name: docxwendangchuli
description: "Use this skill when the user wants to create, read, edit, or manipulate Word documents (.docx files). Triggers include: any mention of 'Word doc', 'word document', '.docx', 'Word文档', '报告', '论文', '通知', '请示', '申请书', '会议纪要', or requests to produce professional documents with formatting like tables of contents, headings, page numbers. Use python-docx for Chinese documents, docx-js for English documents. Also supports extracting/reorganizing content from .docx files, inserting/replacing images, tracked changes, and comments. Batch operations: '合并文档'/'merge docx', '批量转换'/'batch convert' for merging multiple .docx or converting between formats (docx/pdf/md/txt). Do NOT use for PDFs, spreadsheets, Google Docs, or general coding tasks."
license: Proprietary. LICENSE.txt has complete terms
---

# DOCX creation, editing, and analysis

## Overview

A .docx file is a ZIP archive containing XML files. 

This skill provides three document creation paths:
- **Chinese documents**: `python-docx` with pre-built templates (通知, 报告, 论文, etc.)
- **English documents**: `docx-js` (Node.js)
- **Editing existing**: Unpack → Edit XML → Repack

## 写作风格：讲人话，降 AI 味

**以下规则适用于所有生成文档内容的场景（正文段落、摘要、报告、通知等），生成内容时必须遵守。**

### 用词

- **用大白话** — 写"大家商量之后定了下面几条规矩"，不写"经充分研讨与协商，现制定如下规范"
- **砍掉套话** — 删掉"综上所述""值得注意的是""毋庸置疑""显而易见""众所周知"
- **禁用虚词堆砌** — 不写"进行了深入的探讨和研究"，写"聊了聊这事怎么办"或直接说结论
- **少用"进行/作出/予以/加以"** — "进行了修改"→"改了"，"作出说明"→"说明一下"，"予以考虑"→"考虑"
- **具体比抽象好** — 不写"取得显著成效"，写具体数字或事实；没有就别说
- **用短词不用长词** — "为了"不写"为了进一步"，"现在"不写"当前"，"因为"不写"鉴于"

### 句式

- **长短交错** — 别连着三个长句，也别全是短句。一句长一句短，读着不累
- **主动不被动** — "我们检查了设备"，不写"设备被相关人员进行了检查"
- **能拆就拆** — 一个句子只说一件事，超过 40 字的句子考虑拆开
- **别用"首先/其次/再次/最后"开火车** — 换着来。可以列点，也可以用"第一""第二"，或者干脆什么都不加，用自然过渡

### 内容

- **说人关心的事** — 通知先说影响谁、要干嘛；报告先说结论再说过程
- **砍空话** — "为进一步加强和规范……根据……文件精神，结合……实际情况" → 直接说要干嘛
- **加具体细节** — "请大家注意安全"改成"出实验室之前关掉电源、锁好门窗"
- **别过度总结** — 结尾不用再概括一遍前面说过的，说完就停

### 检查清单

生成完内容后过一遍：
1. 能读出声来吗？哪里拗口改哪里
2. 有废话吗？删掉不影响意思的就删
3. 换成口头跟同事说的方式，是不是更短更好懂？
4. 有没有"首先其次再次最后"扎堆？有就拆开

> 注意：排版格式（标题层级、字体、页边距、发文字号等）仍按各模板的规范走，这里的规则只管**内容文字本身**。

---

## Quick Reference

| Task | Approach |
|------|----------|
| Read/analyze content | `pandoc` or unpack for raw XML |
| Batch merge documents | `python scripts/batch_merge.py` — see Batch Processing |
| Batch format convert | `python scripts/batch_convert.py` — see Batch Processing |
| Create Chinese document | Use `python-docx` templates - see Chinese Document Generation |
| Create English document | Use `docx-js` - see Creating New Documents |
| Edit existing document | Unpack → edit XML → repack - see Editing Existing Documents |

### Converting .doc to .docx

Legacy `.doc` files must be converted before editing:

```bash
python scripts/office/soffice.py --headless --convert-to docx document.doc
```

### Reading Content

```bash
# Text extraction with tracked changes
pandoc --track-changes=all document.docx -o output.md

# Raw XML access
python scripts/office/unpack.py document.docx unpacked/
```

### Converting to Images

```bash
python scripts/office/soffice.py --headless --convert-to pdf document.docx
pdftoppm -jpeg -r 150 document.pdf page
```

### Accepting Tracked Changes

To produce a clean document with all tracked changes accepted (requires LibreOffice):

```bash
python scripts/accept_changes.py input.docx output.docx
```

---

## Creating New Documents

Generate .docx files with JavaScript, then validate. Install: `npm install -g docx`

### Setup
```javascript
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
        Header, Footer, AlignmentType, PageOrientation, LevelFormat, ExternalHyperlink,
        InternalHyperlink, Bookmark, FootnoteReferenceRun, PositionalTab,
        PositionalTabAlignment, PositionalTabRelativeTo, PositionalTabLeader,
        TabStopType, TabStopPosition, Column, SectionType,
        TableOfContents, HeadingLevel, BorderStyle, WidthType, ShadingType,
        VerticalAlign, PageNumber, PageBreak } = require('docx');

const doc = new Document({ sections: [{ children: [/* content */] }] });
Packer.toBuffer(doc).then(buffer => fs.writeFileSync("doc.docx", buffer));
```

### Validation
After creating the file, validate it. If validation fails, unpack, fix the XML, and repack.
```bash
python scripts/office/validate.py doc.docx
```

### Page Size

```javascript
// CRITICAL: docx-js defaults to A4, not US Letter
// Always set page size explicitly for consistent results
sections: [{
  properties: {
    page: {
      size: {
        width: 12240,   // 8.5 inches in DXA
        height: 15840   // 11 inches in DXA
      },
      margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } // 1 inch margins
    }
  },
  children: [/* content */]
}]
```

**Common page sizes (DXA units, 1440 DXA = 1 inch):**

| Paper | Width | Height | Content Width (1" margins) |
|-------|-------|--------|---------------------------|
| US Letter | 12,240 | 15,840 | 9,360 |
| A4 (default) | 11,906 | 16,838 | 9,026 |

**Landscape orientation:** docx-js swaps width/height internally, so pass portrait dimensions and let it handle the swap:
```javascript
size: {
  width: 12240,   // Pass SHORT edge as width
  height: 15840,  // Pass LONG edge as height
  orientation: PageOrientation.LANDSCAPE  // docx-js swaps them in the XML
},
// Content width = 15840 - left margin - right margin (uses the long edge)
```

### Styles (Override Built-in Headings)

Use Arial as the default font (universally supported). Keep titles black for readability.

```javascript
const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 24 } } }, // 12pt default
    paragraphStyles: [
      // IMPORTANT: Use exact IDs to override built-in styles
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial" },
        paragraph: { spacing: { before: 240, after: 240 }, outlineLevel: 0 } }, // outlineLevel required for TOC
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial" },
        paragraph: { spacing: { before: 180, after: 180 }, outlineLevel: 1 } },
    ]
  },
  sections: [{
    children: [
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Title")] }),
    ]
  }]
});
```

### Lists (NEVER use unicode bullets)

```javascript
// ❌ WRONG - never manually insert bullet characters
new Paragraph({ children: [new TextRun("• Item")] })  // BAD
new Paragraph({ children: [new TextRun("\u2022 Item")] })  // BAD

// ✅ CORRECT - use numbering config with LevelFormat.BULLET
const doc = new Document({
  numbering: {
    config: [
      { reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "numbers",
        levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ]
  },
  sections: [{
    children: [
      new Paragraph({ numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Bullet item")] }),
      new Paragraph({ numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("Numbered item")] }),
    ]
  }]
});

// ⚠️ Each reference creates INDEPENDENT numbering
// Same reference = continues (1,2,3 then 4,5,6)
// Different reference = restarts (1,2,3 then 1,2,3)
```

### Tables

**CRITICAL: Tables need dual widths** - set both `columnWidths` on the table AND `width` on each cell. Without both, tables render incorrectly on some platforms.

```javascript
// CRITICAL: Always set table width for consistent rendering
// CRITICAL: Use ShadingType.CLEAR (not SOLID) to prevent black backgrounds
const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };

new Table({
  width: { size: 9360, type: WidthType.DXA }, // Always use DXA (percentages break in Google Docs)
  columnWidths: [4680, 4680], // Must sum to table width (DXA: 1440 = 1 inch)
  rows: [
    new TableRow({
      children: [
        new TableCell({
          borders,
          width: { size: 4680, type: WidthType.DXA }, // Also set on each cell
          shading: { fill: "D5E8F0", type: ShadingType.CLEAR }, // CLEAR not SOLID
          margins: { top: 80, bottom: 80, left: 120, right: 120 }, // Cell padding (internal, not added to width)
          children: [new Paragraph({ children: [new TextRun("Cell")] })]
        })
      ]
    })
  ]
})
```

**Table width calculation:**

Always use `WidthType.DXA` — `WidthType.PERCENTAGE` breaks in Google Docs.

```javascript
// Table width = sum of columnWidths = content width
// US Letter with 1" margins: 12240 - 2880 = 9360 DXA
width: { size: 9360, type: WidthType.DXA },
columnWidths: [7000, 2360]  // Must sum to table width
```

**Width rules:**
- **Always use `WidthType.DXA`** — never `WidthType.PERCENTAGE` (incompatible with Google Docs)
- Table width must equal the sum of `columnWidths`
- Cell `width` must match corresponding `columnWidth`
- Cell `margins` are internal padding - they reduce content area, not add to cell width
- For full-width tables: use content width (page width minus left and right margins)

### Images

```javascript
// CRITICAL: type parameter is REQUIRED
new Paragraph({
  children: [new ImageRun({
    type: "png", // Required: png, jpg, jpeg, gif, bmp, svg
    data: fs.readFileSync("image.png"),
    transformation: { width: 200, height: 150 },
    altText: { title: "Title", description: "Desc", name: "Name" } // All three required
  })]
})
```

### Page Breaks

```javascript
// CRITICAL: PageBreak must be inside a Paragraph
new Paragraph({ children: [new PageBreak()] })

// Or use pageBreakBefore
new Paragraph({ pageBreakBefore: true, children: [new TextRun("New page")] })
```

### Hyperlinks

```javascript
// External link
new Paragraph({
  children: [new ExternalHyperlink({
    children: [new TextRun({ text: "Click here", style: "Hyperlink" })],
    link: "https://example.com",
  })]
})

// Internal link (bookmark + reference)
// 1. Create bookmark at destination
new Paragraph({ heading: HeadingLevel.HEADING_1, children: [
  new Bookmark({ id: "chapter1", children: [new TextRun("Chapter 1")] }),
]})
// 2. Link to it
new Paragraph({ children: [new InternalHyperlink({
  children: [new TextRun({ text: "See Chapter 1", style: "Hyperlink" })],
  anchor: "chapter1",
})]})
```

### Footnotes

```javascript
const doc = new Document({
  footnotes: {
    1: { children: [new Paragraph("Source: Annual Report 2024")] },
    2: { children: [new Paragraph("See appendix for methodology")] },
  },
  sections: [{
    children: [new Paragraph({
      children: [
        new TextRun("Revenue grew 15%"),
        new FootnoteReferenceRun(1),
        new TextRun(" using adjusted metrics"),
        new FootnoteReferenceRun(2),
      ],
    })]
  }]
});
```

### Tab Stops

```javascript
// Right-align text on same line (e.g., date opposite a title)
new Paragraph({
  children: [
    new TextRun("Company Name"),
    new TextRun("\tJanuary 2025"),
  ],
  tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
})

// Dot leader (e.g., TOC-style)
new Paragraph({
  children: [
    new TextRun("Introduction"),
    new TextRun({ children: [
      new PositionalTab({
        alignment: PositionalTabAlignment.RIGHT,
        relativeTo: PositionalTabRelativeTo.MARGIN,
        leader: PositionalTabLeader.DOT,
      }),
      "3",
    ]}),
  ],
})
```

### Multi-Column Layouts

```javascript
// Equal-width columns
sections: [{
  properties: {
    column: {
      count: 2,          // number of columns
      space: 720,        // gap between columns in DXA (720 = 0.5 inch)
      equalWidth: true,
      separate: true,    // vertical line between columns
    },
  },
  children: [/* content flows naturally across columns */]
}]

// Custom-width columns (equalWidth must be false)
sections: [{
  properties: {
    column: {
      equalWidth: false,
      children: [
        new Column({ width: 5400, space: 720 }),
        new Column({ width: 3240 }),
      ],
    },
  },
  children: [/* content */]
}]
```

Force a column break with a new section using `type: SectionType.NEXT_COLUMN`.

### Table of Contents

```javascript
// CRITICAL: Headings must use HeadingLevel ONLY - no custom styles
new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-3" })
```

### Headers/Footers

```javascript
sections: [{
  properties: {
    page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } // 1440 = 1 inch
  },
  headers: {
    default: new Header({ children: [new Paragraph({ children: [new TextRun("Header")] })] })
  },
  footers: {
    default: new Footer({ children: [new Paragraph({
      children: [new TextRun("Page "), new TextRun({ children: [PageNumber.CURRENT] })]
    })] })
  },
  children: [/* content */]
}]
```

### Critical Rules for docx-js

- **Set page size explicitly** - docx-js defaults to A4; use US Letter (12240 x 15840 DXA) for US documents
- **Landscape: pass portrait dimensions** - docx-js swaps width/height internally; pass short edge as `width`, long edge as `height`, and set `orientation: PageOrientation.LANDSCAPE`
- **Never use `\n`** - use separate Paragraph elements
- **Never use unicode bullets** - use `LevelFormat.BULLET` with numbering config
- **PageBreak must be in Paragraph** - standalone creates invalid XML
- **ImageRun requires `type`** - always specify png/jpg/etc
- **Always set table `width` with DXA** - never use `WidthType.PERCENTAGE` (breaks in Google Docs)
- **Tables need dual widths** - `columnWidths` array AND cell `width`, both must match
- **Table width = sum of columnWidths** - for DXA, ensure they add up exactly
- **Always add cell margins** - use `margins: { top: 80, bottom: 80, left: 120, right: 120 }` for readable padding
- **Use `ShadingType.CLEAR`** - never SOLID for table shading
- **Never use tables as dividers/rules** - cells have minimum height and render as empty boxes (including in headers/footers); use `border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "2E75B6", space: 1 } }` on a Paragraph instead. For two-column footers, use tab stops (see Tab Stops section), not tables
- **TOC requires HeadingLevel only** - no custom styles on heading paragraphs
- **Override built-in styles** - use exact IDs: "Heading1", "Heading2", etc.
- **Include `outlineLevel`** - required for TOC (0 for H1, 1 for H2, etc.)

---

## Editing Existing Documents

**Follow all 3 steps in order.**

### Step 1: Unpack
```bash
python scripts/office/unpack.py document.docx unpacked/
```
Extracts XML, pretty-prints, merges adjacent runs, and converts smart quotes to XML entities (`&#x201C;` etc.) so they survive editing. Use `--merge-runs false` to skip run merging.

### Step 2: Edit XML

Edit files in `unpacked/word/`. See XML Reference below for patterns.

**Use "Claude" as the author** for tracked changes and comments, unless the user explicitly requests use of a different name.

**Use the Edit tool directly for string replacement. Do not write Python scripts.** Scripts introduce unnecessary complexity. The Edit tool shows exactly what is being replaced.

**CRITICAL: Use smart quotes for new content.** When adding text with apostrophes or quotes, use XML entities to produce smart quotes:
```xml
<!-- Use these entities for professional typography -->
<w:t>Here&#x2019;s a quote: &#x201C;Hello&#x201D;</w:t>
```
| Entity | Character |
|--------|-----------|
| `&#x2018;` | ‘ (left single) |
| `&#x2019;` | ’ (right single / apostrophe) |
| `&#x201C;` | “ (left double) |
| `&#x201D;` | ” (right double) |

**Adding comments:** Use `comment.py` to handle boilerplate across multiple XML files (text must be pre-escaped XML):
```bash
python scripts/comment.py unpacked/ 0 "Comment text with &amp; and &#x2019;"
python scripts/comment.py unpacked/ 1 "Reply text" --parent 0  # reply to comment 0
python scripts/comment.py unpacked/ 0 "Text" --author "Custom Author"  # custom author name
```
Then add markers to document.xml (see Comments in XML Reference).

### Step 3: Pack
```bash
python scripts/office/pack.py unpacked/ output.docx --original document.docx
```
Validates with auto-repair, condenses XML, and creates DOCX. Use `--validate false` to skip.

**Auto-repair will fix:**
- `durableId` >= 0x7FFFFFFF (regenerates valid ID)
- Missing `xml:space="preserve"` on `<w:t>` with whitespace

**Auto-repair won't fix:**
- Malformed XML, invalid element nesting, missing relationships, schema violations

### Common Pitfalls

- **Replace entire `<w:r>` elements**: When adding tracked changes, replace the whole `<w:r>...</w:r>` block with `<w:del>...<w:ins>...` as siblings. Don't inject tracked change tags inside a run.
- **Preserve `<w:rPr>` formatting**: Copy the original run's `<w:rPr>` block into your tracked change runs to maintain bold, font size, etc.

---

## XML Reference

### Schema Compliance

- **Element order in `<w:pPr>`**: `<w:pStyle>`, `<w:numPr>`, `<w:spacing>`, `<w:ind>`, `<w:jc>`, `<w:rPr>` last
- **Whitespace**: Add `xml:space="preserve"` to `<w:t>` with leading/trailing spaces
- **RSIDs**: Must be 8-digit hex (e.g., `00AB1234`)

### Tracked Changes

**Insertion:**
```xml
<w:ins w:id="1" w:author="Claude" w:date="2025-01-01T00:00:00Z">
  <w:r><w:t>inserted text</w:t></w:r>
</w:ins>
```

**Deletion:**
```xml
<w:del w:id="2" w:author="Claude" w:date="2025-01-01T00:00:00Z">
  <w:r><w:delText>deleted text</w:delText></w:r>
</w:del>
```

**Inside `<w:del>`**: Use `<w:delText>` instead of `<w:t>`, and `<w:delInstrText>` instead of `<w:instrText>`.

**Minimal edits** - only mark what changes:
```xml
<!-- Change "30 days" to "60 days" -->
<w:r><w:t>The term is </w:t></w:r>
<w:del w:id="1" w:author="Claude" w:date="...">
  <w:r><w:delText>30</w:delText></w:r>
</w:del>
<w:ins w:id="2" w:author="Claude" w:date="...">
  <w:r><w:t>60</w:t></w:r>
</w:ins>
<w:r><w:t> days.</w:t></w:r>
```

**Deleting entire paragraphs/list items** - when removing ALL content from a paragraph, also mark the paragraph mark as deleted so it merges with the next paragraph. Add `<w:del/>` inside `<w:pPr><w:rPr>`:
```xml
<w:p>
  <w:pPr>
    <w:numPr>...</w:numPr>  <!-- list numbering if present -->
    <w:rPr>
      <w:del w:id="1" w:author="Claude" w:date="2025-01-01T00:00:00Z"/>
    </w:rPr>
  </w:pPr>
  <w:del w:id="2" w:author="Claude" w:date="2025-01-01T00:00:00Z">
    <w:r><w:delText>Entire paragraph content being deleted...</w:delText></w:r>
  </w:del>
</w:p>
```
Without the `<w:del/>` in `<w:pPr><w:rPr>`, accepting changes leaves an empty paragraph/list item.

**Rejecting another author's insertion** - nest deletion inside their insertion:
```xml
<w:ins w:author="Jane" w:id="5">
  <w:del w:author="Claude" w:id="10">
    <w:r><w:delText>their inserted text</w:delText></w:r>
  </w:del>
</w:ins>
```

**Restoring another author's deletion** - add insertion after (don't modify their deletion):
```xml
<w:del w:author="Jane" w:id="5">
  <w:r><w:delText>deleted text</w:delText></w:r>
</w:del>
<w:ins w:author="Claude" w:id="10">
  <w:r><w:t>deleted text</w:t></w:r>
</w:ins>
```

### Comments

After running `comment.py` (see Step 2), add markers to document.xml. For replies, use `--parent` flag and nest markers inside the parent's.

**CRITICAL: `<w:commentRangeStart>` and `<w:commentRangeEnd>` are siblings of `<w:r>`, never inside `<w:r>`.**

```xml
<!-- Comment markers are direct children of w:p, never inside w:r -->
<w:commentRangeStart w:id="0"/>
<w:del w:id="1" w:author="Claude" w:date="2025-01-01T00:00:00Z">
  <w:r><w:delText>deleted</w:delText></w:r>
</w:del>
<w:r><w:t> more text</w:t></w:r>
<w:commentRangeEnd w:id="0"/>
<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>

<!-- Comment 0 with reply 1 nested inside -->
<w:commentRangeStart w:id="0"/>
  <w:commentRangeStart w:id="1"/>
  <w:r><w:t>text</w:t></w:r>
  <w:commentRangeEnd w:id="1"/>
<w:commentRangeEnd w:id="0"/>
<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>
<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="1"/></w:r>
```

### Images

1. Add image file to `word/media/`
2. Add relationship to `word/_rels/document.xml.rels`:
```xml
<Relationship Id="rId5" Type=".../image" Target="media/image1.png"/>
```
3. Add content type to `[Content_Types].xml`:
```xml
<Default Extension="png" ContentType="image/png"/>
```
4. Reference in document.xml:
```xml
<w:drawing>
  <wp:inline>
    <wp:extent cx="914400" cy="914400"/>  <!-- EMUs: 914400 = 1 inch -->
    <a:graphic>
      <a:graphicData uri=".../picture">
        <pic:pic>
          <pic:blipFill><a:blip r:embed="rId5"/></pic:blipFill>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
  </wp:inline>
</w:drawing>
```

---

## Chinese Document Generation (中文文档生成)

Use `python-docx` + `scripts/chinese_format.py` for Chinese documents. This gives native CJK font support, standard Chinese formatting, and pre-built templates.

### Setup

```bash
pip install python-docx lxml
```

### Quick Reference: Templates

All templates use a consistent calling convention:

```python
from scripts.docx_gen.official import generate_tongzhi, generate_qingshi, generate_baogao, generate_han
from scripts.docx_gen.academic import generate_course_paper, generate_thesis, generate_proposal
from scripts.docx_gen.general import generate_application, generate_meeting_minutes
```

| Category | Document | Function | Key Features |
|----------|----------|----------|--------------|
| 党政公文 | 通知 (Notice) | `generate_tongzhi()` | Official margins, red header, 发文字号, 印章 |
| 党政公文 | 请示 (Request) | `generate_qingshi()` | 缘由→事项→结语 structure |
| 党政公文 | 报告 (Report) | `generate_baogao()` | 前言→主体章节→结语 |
| 党政公文 | 函 (Letter) | `generate_han()` | Sender/receiver, contact info |
| 学术论文 | 课程论文 (Course Paper) | `generate_course_paper()` | Abstract, keywords, sections, references |
| 学术论文 | 学位论文 (Thesis) | `generate_thesis()` | Cover, CN+EN abstract, chapters, acknowledgments |
| 学术论文 | 开题报告 (Proposal) | `generate_proposal()` | Background, lit review, methods, schedule |
| 日常文书 | 申请书 (Application) | `generate_application()` | 称呼→正文→此致/敬礼→署名 |
| 日常文书 | 会议纪要 (Minutes) | `generate_meeting_minutes()` | 会议信息→议题→决议→记录人 |

### Chinese Formatting Presets

**Font Hierarchy (黑体/宋体 Rule):**
- Headings: 黑体 (Heiti) + Times New Roman
- Body: 宋体 (Songti) + Times New Roman
- Captions: 楷体 (Kaiti) + Arial
- Official body: 仿宋 (FangSong) + Times New Roman

**Font Size Presets** (from `chinese_format.FontPreset`):

| Preset | Chinese | Western | Size | Bold |
|--------|---------|---------|------|------|
| TITLE_LARGE | 黑体 | Times New Roman | 26pt (一号) | Yes |
| TITLE | 黑体 | Times New Roman | 22pt (二号) | Yes |
| H1 | 黑体 | Times New Roman | 16pt (三号) | Yes |
| H2 | 黑体 | Times New Roman | 15pt (小三) | Yes |
| H3 | 黑体 | Times New Roman | 12pt (小四) | Yes |
| BODY | 宋体 | Times New Roman | 12pt (小四) | No |
| BODY_SMALL | 宋体 | Times New Roman | 10.5pt (五号) | No |
| CAPTION | 楷体 | Arial | 10.5pt (五号) | No |
| OFFICIAL_BODY | 仿宋 | Times New Roman | 16pt (三号) | No |

**Page Settings** (from `chinese_format`):

| Preset | Top | Bottom | Left | Right | Use Case |
|--------|-----|--------|------|-------|----------|
| PAGE_DEFAULT | 2.54cm | 2.54cm | 3.18cm | 3.18cm | General documents |
| PAGE_OFFICIAL | 3.7cm | 3.5cm | 2.8cm | 2.6cm | Government docs (GB/T 9704) |
| PAGE_ACADEMIC | 2.5cm | 2.5cm | 2.5cm | 2.5cm | Academic papers (GB/T 7713) |
| PAGE_REPORT | 3.0cm | 2.5cm | 2.8cm | 2.6cm | Work reports |

**Paragraph Presets:**

| Preset | Line Spacing | Indent | Alignment |
|--------|-------------|--------|-----------|
| PARA_BODY | 1.25x | 2 chars | Justified |
| PARA_TITLE | 1.5x | 0 | Centered |
| PARA_HEADING | 1.25x | 0 | Left |
| PARA_CAPTION | 1.0x | 0 | Centered |
| PARA_OFFICIAL | 28pt | 2 chars | Justified |
| PARA_SIGNATURE | 1.5x | 0 | Right |

### Manual Formatting with python-docx

For custom documents not covered by templates, use the `ChineseDocument` factory:

```python
from scripts.docx_gen.base import ChineseDocument
from scripts.chinese_format import PAGE_DEFAULT, FontPreset

doc = ChineseDocument.create(PAGE_DEFAULT)
ChineseDocument.add_title(doc, "Document Title")
ChineseDocument.add_heading(doc, "Section 1", level=1)
ChineseDocument.add_body(doc, "Body paragraph text...")
ChineseDocument.add_caption(doc, "Table 1: Data Summary")
ChineseDocument.add_signature(doc, "Author Name")
ChineseDocument.save(doc, "output.docx")
```

For mixed Chinese/English runs, use `set_run_font()` directly:

```python
from scripts.chinese_format import set_run_font
p = doc.add_paragraph()
run_cn = p.add_run("中文内容")
set_run_font(run_cn, "宋体", "Times New Roman", 12)
run_en = p.add_run("English text")
set_run_font(run_en, "Times New Roman", "Times New Roman", 12)
```

### Critical Rules for Chinese Documents

- **Always set both CN and EN fonts** — use `set_run_font(cn="黑体", en="Times New Roman", size_pt=16)` not `run.font.name`
- **Never use spaces for indentation** — use `paragraph_format.first_line_indent = Pt(24)` (2 chars × 12pt)
- **Use `ParagraphPreset` for consistent spacing** — don't set line_spacing/space_before/space_after manually
- **Official documents (党政公文) use 仿宋 body** — not 宋体; academic/general documents use 宋体
- **Chinese fonts are pre-installed on all Chinese Windows systems** — 宋体, 黑体, 楷体, 仿宋 require no download
- **方正小标宋简体 and 仿宋_GB2312 are NOT built-in** — fall back to 黑体 and 仿宋 respectively

---

## Batch Processing（批量处理）

### 批量合并文档

将多个 .docx 文件合并为一个文件：

```bash
# 合并指定的文件
python scripts/batch_merge.py file1.docx file2.docx file3.docx -o merged.docx

# 使用通配符合并
python scripts/batch_merge.py chapter*.docx -o book.docx

# 无分页符合并（内容连续排列）
python scripts/batch_merge.py *.docx -o combined.docx --separator none
```

### 批量格式转换

将多个文档从一种格式批量转换为另一种：

```bash
# docx → PDF（需要 LibreOffice）
python scripts/batch_convert.py --from docx --to pdf report1.docx report2.docx -o ./pdfs/
python scripts/batch_convert.py --from docx --to pdf ./documents/ -o ./pdfs/

# doc → docx（需要 LibreOffice）
python scripts/batch_convert.py --from doc --to docx *.doc -o ./converted/

# docx → Markdown（需要 pandoc）
python scripts/batch_convert.py --from docx --to md *.docx -o ./markdown/

# docx → 纯文本（需要 pandoc）
python scripts/batch_convert.py --from docx --to txt *.docx -o ./text/
```

### 批量操作参数速查

| 脚本 | 功能 | 依赖 |
|------|------|------|
| `batch_merge.py` | 合并多个 docx → 一个文件 | python-docx |
| `batch_convert.py --from docx --to pdf` | docx 批量转 PDF | LibreOffice |
| `batch_convert.py --from doc --to docx` | doc 批量转 docx | LibreOffice |
| `batch_convert.py --from docx --to md` | docx 批量转 Markdown | pandoc |
| `batch_convert.py --from docx --to txt` | docx 批量转纯文本 | pandoc |

### 注意事项

- **通配符支持** — 两个脚本都支持 `*.docx` 通配符和目录输入
- **排序** — 文件按名称字母序处理，要控制合并顺序请逐个指定文件名
- **合并保留样式** — `batch_merge.py` 以第一个文档为基础模板，后续文档的内容直接追加
- **转换超时** — 默认单文件超时 300 秒，大文件可用 `--timeout 600` 加长

---

## Dependencies

- **pandoc**: Text extraction + batch docx→md/txt conversion
- **docx**: `npm install -g docx` (new English documents)
- **python-docx**: `pip install python-docx lxml` (new Chinese documents + batch merge)
- **LibreOffice**: PDF conversion + batch format conversion (auto-configured for sandboxed environments via `scripts/office/soffice.py`)
- **Poppler**: `pdftoppm` for images

---
name: gongwen-format
description: >
  Chinese government document (党政机关公文) formatter — converts markdown or plain-text
  drafts to .docx files that conform to GB/T 9704-2012 formatting standards.
  TRIGGERS: 公文格式, 政府公文, 红头文件, 党政机关文档, 正式公文排版,
  GB/T 9704, A4排版, 28字22行, 方正小标宋, 仿宋三号.
---

# Gongwen Format — 党政机关公文排版

## Overview

This skill converts a draft document (Markdown, plain text, or structured data)
into a properly formatted **Chinese government official document (.docx)** following
the **党政机关公文格式 (GB/T 9704-2012)** standard.

Use `scripts/generate.py` as a starting point — copy it to the working directory,
adapt the `build_content(doc)` function with the user's actual document content,
then run it.

## When to Use

- User asks to format a document as 公文 / 政府公文 / 红头文件
- User provides detailed formatting specs matching GB/T 9704
- User mentions: 公文格式, 党政机关公文, A4排版, 28字×22行, 方正小标宋
- User wants proper Chinese document margins, fonts, and line spacing

## Formatting Cheat Sheet

### Page Setup
| Property | Value | DXA equivalent |
|----------|-------|----------------|
| Paper | A4 (210mm × 297mm) | w=11906, h=16838 |
| Top margin | 37mm | ~2098 |
| Bottom margin | 35mm | ~1984 |
| Left margin | 28mm | ~1587 |
| Right margin | 26mm | ~1474 |
| Text area (版心) | 156mm × 225mm | — |

### Fonts and Sizes
| Element | Chinese Font | Size | Bold |
|---------|-------------|------|------|
| Document title (标题) | 方正小标宋简体 | 2号 (22pt) | no |
| Body text (正文) | 仿宋_GB2312 | 3号 (16pt) | no |
| Level-1 heading (一级标题) | 黑体 | 3号 (16pt) | yes |
| Level-2 heading (二级标题) | 楷体_GB2312 | 3号 (16pt) | yes |
| Level-3/4 heading (三四级) | 仿宋_GB2312 | 3号 (16pt) | yes |
| Numbers / Latin text | Times New Roman | same as context | — |

### Fallback Chain
When primary fonts are unavailable, Word auto-falls back to:
- `方正小标宋简体` → `宋体` (SimSun)
- `仿宋_GB2312` → `仿宋` (FangSong)
- `楷体_GB2312` → `楷体` (KaiTi)
- `黑体` → hardcoded, always available (SimHei)

### Line Spacing
- **Fixed 28.9pt** for all body text, headings, and empty lines
- Code blocks: fixed 18pt for readability
- This yields ~22 lines per page at 3号 body size

### First-Line Indent
- **2 characters** = 32pt = ~640 twips (at 3号 = 16pt × 2)
- Applied to all body paragraphs, NOT to headings or centered lines

### Heading Numbering
| Level | Sequence | Font |
|-------|----------|------|
| 一级 (h1) | 一、二、三、… | 黑体 |
| 二级 (h2) | （一）（二）（三）… | 楷体_GB2312 |
| 三级 (h3) | 1. 2. 3. … | 仿宋_GB2312 |
| 四级 (h4) | （1）（2）（3）… | 仿宋_GB2312 |

### Title Placement
- Located below red separator line, with **2 blank lines** gap
- Centered, 2号 (22pt) 方正小标宋简体
- Multi-line titles: trapezoid or diamond arrangement

### 主送机关 (Addressee)
- 1 blank line below title
- Flush-left (顶格), 3号 仿宋_GB2312
- Multiple addresses separated by Chinese顿号 (、)
- Ends with full-width colon (：)

### Tables
- Header row: gray background (#E8EDF5), bold 黑体 text
- Border: single-line, #999999
- Cell padding: ~40/80 DXA

---

## Workflow

### Step 1: Gather Content
Read the source document (Markdown, text, or user-provided content).
Identify:
- Title line
- Metadata / subtitle lines
- Section headings and their levels
- Body paragraphs
- Tables, code blocks, lists

### Step 2: Map Content to docx Elements
Use `python-docx` (installed via `pip install python-docx`) with the helper
functions from `scripts/generate.py`:

```python
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING

# Page setup
section = doc.sections[0]
section.page_width = Cm(21.0)
section.page_height = Cm(29.7)
section.top_margin = Cm(3.7)
section.bottom_margin = Cm(3.5)
section.left_margin = Cm(2.8)
section.right_margin = Cm(2.6)

# Set east-Asian font via XML
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

def set_font(run, east_asian, ascii_font="Times New Roman", bold=False, size=Pt(16)):
    run.font.size = size
    run.bold = bold
    run.font.name = ascii_font
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} />')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), east_asian)
    rFonts.set(qn('w:ascii'), ascii_font)
    rFonts.set(qn('w:hAnsi'), ascii_font)
```

### Step 3: Generate and Validate
```bash
python generate.py
```

The output .docx can be opened directly in Microsoft Word or WPS Office.
Font fallback ensures correct display even when primary fonts are absent.

---

## Key Paragraph Patterns

```python
# Line spacing (fixed 28.9pt)
pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
pf.line_spacing = Pt(28.9)

# First-line indent (2 chars = 32pt)
pf.first_line_indent = Pt(32)

# Title — centered, 22pt, 方正小标宋简体
set_font(run, '方正小标宋简体', size=Pt(22))
pf.alignment = WD_ALIGN_PARAGRAPH.CENTER

# H1 — 黑体 16pt bold
set_font(run, '黑体', bold=True)

# H2 — 楷体_GB2312 16pt bold
set_font(run, '楷体_GB2312', bold=True)

# Body — 仿宋_GB2312 16pt, first-line indent
set_font(run, '仿宋_GB2312')
pf.first_line_indent = Pt(32)
```

---

## Dependencies

- Python 3.8+
- `python-docx` (`pip install python-docx`)

## Reference

- GB/T 9704-2012 党政机关公文格式
- GB/T 9704-1999 (superseded)

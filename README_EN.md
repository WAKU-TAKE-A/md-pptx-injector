# md-pptx-injector

A Python script that converts Markdown files into PowerPoint presentations (.pptx).  
It injects Markdown content into template layouts and placeholders, giving you full control over slide structure.

---

## Requirements

- Python 3.10 or later
- python-pptx
- Pillow (required for image insertion)

```bash
pip install python-pptx Pillow
```

---

## Usage

```bash
python md-pptx-injector.py src.md dst.pptx --template ref.pptx
python md-pptx-injector.py src.md dst.pptx --template ref.pptx -v
```

| Argument | Description |
|----------|-------------|
| `src.md` | Input Markdown file |
| `dst.pptx` | Output PowerPoint file |
| `--template ref.pptx` | Template PPTX file (default: `template.pptx`) |
| `-v` | Verbose logging for debugging |

### Path Resolution Rules

- Absolute path → used as-is
- Starts with `./` or `.\` → relative to current working directory
- Otherwise → relative to the script (or exe) directory

---

## Markdown Syntax

### Page Breaks

A new slide is created by any of the following:

| Syntax | Behavior |
|--------|----------|
| `# Heading` / `## Heading` / `### Heading` | Page break |
| `<!-- new_page -->` | Unconditional page break |
| `<!-- new_page="LayoutName" -->` | Page break with specified layout |
| `<!-- layout="LayoutName" -->` immediately followed by `#/##/###` | Page break (layout tag must directly precede the heading) |

> **Notes:**
> - `---` (horizontal rule) is ignored everywhere except YAML front matter.
> - When YAML front matter is present, a level-1 heading (`#`) is kept on the same page and does not create an empty title slide.

### YAML Front Matter

Place at the very beginning of the document.

```markdown
---
title: Presentation Title
subtitle: Subtitle Text
author: Author Name
toc: true
toc_title: Table of Contents
indent: 2
font_size_l0: 18
font_size_l1: 16
font_size_l2: 14
font_size_l3: 12
font_size_l4: 10
---
```

| Key | Description |
|-----|-------------|
| `title` | Title slide title |
| `subtitle` | Title slide subtitle |
| `author` | Author name (appended to subtitle) |
| `toc: true` | Generate a TOC slide at the end |
| `toc_title` | TOC slide title (default: `Table of Contents`) |
| `indent` | List indent width in spaces (default: 2) |
| `font_size_l0` – `font_size_l4` | Font size (pt) per level, applied document-wide |

### Layout Override

```markdown
<!-- layout="sample02" -->
### Page Title
```

- The `<!-- layout="..." -->` tag is ignored if the **very next line** is not a `#/##/###` heading (a warning is shown with `-v`)
- The layout name must match a slide layout name in the template PPTX

Font sizes can also be specified at the layout level (effective for that page only):

```markdown
<!-- layout="sample02" font_size_l0=16 font_size_l1=12 -->
### Page Title
```

### Placeholder Targeting

Target a named shape in the template to inject content into it.

```markdown
<!-- placeholder="holder01" -->
#### ■ Content 1
* Item 1
* Item 2
  * Item 2-1

<!-- placeholder="holder02" -->
#### ■ Content 2
Body text here.
```

Font sizes can also be set per placeholder (effective within that placeholder only):

```markdown
<!-- placeholder="holder01" font_size_l0=14 font_size_l2=10 -->
#### ■ Content
* Item
```

If a placeholder name is not found on the slide, its content falls back to the body area (rescue).

### Font Size Priority

```
YAML (document) < layout (page) < placeholder (holder)
```

Levels without an explicit override inherit from the master template.

### Inline Formatting

```markdown
<b>Bold</b>
<i>Italic</i>
<u>Underline</u>
<b><i>Bold + Italic</i></b>
```

Mismatched or unclosed tags fall back to plain text (a warning is shown with `-v`).

### In-Slide Headings

`####` and deeper headings do not trigger a page break — they are rendered as styled paragraphs within the slide.

```markdown
#### ■ Section heading       → level=0 paragraph
##### □ Sub-section          → level=1 paragraph
###### □ Sub-sub-section     → level=2 paragraph
```

### Bullet Lists

```markdown
* Item 1
* Item 2
  * Item 2-1  (2-space indent)
    * Item 2-1-1 (4-space indent)
- Alternative marker
+ Alternative marker 2
1. Numbered list
```

Up to 5 levels (level 0–4) are supported. Indent width is controlled by `indent`.

### Code Blocks

````markdown
```
code goes here
```
````

Rendered as a gray-background textbox placed at the back of the z-order (rendered behind all other shapes).  
Placeholder content appears in front, so the code block acts as a background layer.

### Images

```markdown
<!-- placeholder="photo01" -->
![Caption text](image.png)
```

- Inserted with contain-fit into the position and size of the `photo01` shape
- If alt text is non-empty, it is written to a shape named `photo01_caption` (if it exists)

Image search order:
1. Same directory as the Markdown file
2. Script (or exe) directory
3. Current working directory

### Tables

```markdown
<!-- placeholder="table01" -->
| Col1  | Col2  | Col3  |
|:------|:-----:|------:|
| A     | B     | C     |
```

- `|:---|` left-align / `|:---:|` center / `|---:|` right-align
- Column widths are proportional to the number of dashes in the separator row (e.g. `|--|----:|` → 1:2 ratio, right-aligned)

### Table of Contents (TOC)

Set `toc: true` in YAML to collect all `##` / `###` headings and generate a TOC slide at the end of the presentation.  
Each entry is a hyperlink that jumps to the corresponding slide when clicked in slideshow mode.

```markdown
---
toc: true
toc_title: Contents
---
```

---

## Example

```markdown
---
title: Sample Presentation
subtitle: For Testing
author: John Doe
toc: true
toc_title: Contents
font_size_l0: 18
font_size_l1: 14
---

# Title
subtitle: Subtitle goes here

## Chapter 1

Chapter description text.

<!-- layout="TwoContent" font_size_l0=16 -->
### Page 1-1

<!-- placeholder="holder01" font_size_l1=12 -->
#### ■ Left Content
* Item A
* Item B
  * Item B-1

<!-- placeholder="holder02" -->
#### ■ Right Content
Body text goes here.

<!-- new_page -->

Blank page content.
```

---

## Automatic Layout Detection

When no `<!-- layout -->` or `<!-- new_page="..." -->` is specified, the layout is chosen automatically based on the heading level:

| Heading Level | Applied Layout |
|---------------|---------------|
| `#` or YAML front matter present | `Title Slide` |
| `##` | `Section Header` |
| `###` or other | `Title and Content` |

If the specified layout is not found in the template, the script falls back to `Title and Content`, then to the first available layout.

---

## Version

Current version: **0.9.4.0**

# md-pptx-injector.py

A Markdown-to-PowerPoint converter with template support, custom layouts, placeholder targeting, inline formatting, and automatic table-of-contents generation.

## Table of Contents

- [Quick Start](#quick-start)
- [Features](#features)
- [Installation](#installation)
- [Usage](#usage)
- [Path Resolution](#path-resolution)
- [Markdown Syntax](#markdown-syntax)
- [Layouts and Placeholders](#layouts-and-placeholders)
- [Advanced Features](#advanced-features)
- [Troubleshooting](#troubleshooting)

---

## Quick Start

```bash
# Install dependencies
pip install python-pptx Pillow

# Basic usage
python md-pptx-injector.py input.md output.pptx --template template.pptx

# With verbose debug output
python md-pptx-injector.py input.md output.pptx --template template.pptx -v
```

---

## Features

- Markdown → PowerPoint conversion using a PPTX template
- Custom layout selection via HTML comments
- Named placeholder targeting for precise content placement
- Explicit page breaks via `<!-- new_page -->` comments
- Inline formatting using `<b>`, `<i>`, `<u>` HTML tags (nestable)
- Bullet lists (`-` `*` `+`) and numbered lists (up to 3 levels)
- Code blocks (` ``` ` fences → grey-background textbox)
- Markdown tables → PowerPoint tables (with column width and alignment control)
- Image insertion with aspect-ratio-preserving contain fit and caption support
- YAML front matter for title slide generation
- Automatic table-of-contents slide with inter-slide hyperlinks (`toc: true`)
- PyInstaller (exe packaging) support

---

## Installation

**Requirements**

- Python 3.9 or later

**Dependencies**

```bash
pip install python-pptx Pillow
```

> `Pillow` is required for image insertion. The script runs without it, but images will be silently skipped.

---

## Usage

### Command Line

```bash
python md-pptx-injector.py <src> <dst> [OPTIONS]
```

| Argument / Option | Description |
|---|---|
| `src` | Source Markdown file |
| `dst` | Destination PowerPoint file |
| `--template PATH` | Template .pptx file (default: `template.pptx`) |
| `-v, --verbose` | Print debug log to stdout |

**Examples**

```bash
# Use default template.pptx
python md-pptx-injector.py slides.md presentation.pptx

# Specify a custom template
python md-pptx-injector.py slides.md presentation.pptx --template custom.pptx

# Debug mode
python md-pptx-injector.py slides.md presentation.pptx -v
```

---

## Path Resolution

### Application Base Directory

| Execution mode | Base directory |
|---|---|
| Script (`.py`) | Directory containing `md-pptx-injector.py` |
| Executable (PyInstaller) | Directory containing `md-pptx-injector.exe` |

### Resolution Order for src, dst, and --template

1. **Absolute path** → used as-is
2. **Relative path that exists in the current working directory** → resolved from CWD (higher priority)
3. **Other relative path** → resolved from the application base directory

### Image Path Resolution Order

1. Directory containing the Markdown file (highest priority)
2. Application base directory
3. Current working directory

---

## Markdown Syntax

### Page Breaks (Slide Boundaries)

A new slide is started by any of the following.

**Headings (`#` / `##` / `###`)**

```markdown
## Slide A

Content for slide A

## Slide B

Content for slide B
```

**Explicit page-break comment**

```markdown
<!-- new_page -->
```

A layout name can be included at the same time (see [Layouts and Placeholders](#layouts-and-placeholders)):

```markdown
<!-- new_page="Two Content" -->
```

> **Note**: The `---` inside a YAML front matter block at the top of a page is never treated as a page break.

---

### Text Formatting

Inline formatting is specified with `<b>`, `<i>`, and `<u>` HTML tags. Tags can be nested.

```markdown
<b>Bold</b>
<i>Italic</i>
<u>Underline</u>
<b><i>Bold and italic</i></b>
Mix <b>bold</b> and <i>italic</i> in a single line.
```

> If a tag is left unclosed or the closing order is wrong, the entire line is rendered as plain text. Run with `-v` to see the warning.

---

### Headings

#### Layout-controlling headings (`#` `##` `###`)

| Markdown | Layout selected |
|---|---|
| `# Title` | `Title Slide` |
| `## Section` | `Section Header` |
| `### Content` | `Title and Content` |

The heading text is written into the title placeholder of the new slide.

#### In-slide headings (`####` `#####`)

These are placed inside the content area and set the **base level** for subsequent paragraphs and bullets.

```markdown
#### Level-0 heading
Text at level 0

##### Level-1 heading
Text at level 1
- Bullet at level 1
  - Bullet at level 2
```

| Heading | Base level |
|---|---|
| `####` | 0 |
| `#####` | 1 |

---

### Bullet Lists and Numbered Lists

```markdown
- Level-0 bullet
  - Level 1 (2 spaces indent)
    - Level 2 (4 spaces indent)

* Asterisk marker also works
+ Plus marker also works

1. Numbered list item (number rendered as literal text)
2. Second item
```

- Indent unit: 2 spaces per level (configurable via `indent:` in front matter)
- Maximum level: 2 (excess indentation is clipped to level 2)
- Numbered list items render their numbers as literal text; auto-numbering is not applied

---

### Code Blocks

````markdown
```
def hello():
    print("Hello, World!")
```
````

Rendered as a grey-background textbox in Courier New font on the slide.

---

### Blank Lines

Blank lines in Markdown produce empty paragraphs (vertical spacing) in PowerPoint.

---

## Layouts and Placeholders

### Layout Selection Priority

1. `<!-- layout="..." -->` comment (highest priority)
2. `<!-- new_page="..." -->` comment
3. Heading level (`#` / `##` / `###`)
4. Presence of YAML front matter with no heading → `Title Slide`

If the specified layout does not exist in the template, the script falls back gracefully: `Title and Content` → first available layout. No error is raised.

---

### Layout Override Comment

```markdown
<!-- layout="Two Content" -->
## Slide Title
```

> **Note**: The `<!-- layout="..." -->` comment must be immediately followed by a `#`/`##`/`###` heading. If the next line is not a heading, the comment is ignored and a warning is logged.

---

### Title Slide

**Pattern A: YAML front matter**

```markdown
---
title: Presentation Title
subtitle: Subtitle Text
author: Author Name
toc: true
indent: 2
---
```

**Pattern B: Markdown headings** (overrides Pattern A)

```markdown
# Presentation Title
subtitle: Subtitle Text
author: Author Name
```

`subtitle` and `author` are placed in the subtitle placeholder, separated by a line break.

**Front matter keys**

| Key | Description |
|---|---|
| `title` | Title text |
| `subtitle` | Subtitle text |
| `author` | Author name |
| `toc` | Set to `true` to auto-generate a TOC slide at the end |
| `indent` | Bullet indent width in spaces (default: `2`) |

---

### Custom Placeholders

Write content into a specific named shape in the template.

```markdown
<!-- placeholder="LeftBox" -->
This content goes into the shape named "LeftBox".
(Captured until a blank line)

<!-- placeholder="RightBox" -->
This content goes into "RightBox".

<!-- placeholder="LeftBox" -->
A second block for the same placeholder is appended with a blank-line separator.
```

**Rules**
- Content is captured until the next blank line.
- Multiple blocks targeting the same placeholder are **appended** (never overwritten).
- If the named placeholder is not found in the slide, the content falls through to rescue (see below).

---

### Rescue Content

Text that has no placeholder assignment is automatically **appended** to the body placeholder.

```markdown
<!-- layout="Two Content" -->
## Slide Title

<!-- placeholder="LeftBox" -->
Content for the left box.

This unassigned text is rescued into the body placeholder.
```

**Rescue conditions**
- Only applies to non-Title-Slide layouts.
- Skipped if the rescue content contains only blank lines.
- Shapes already targeted by an explicit placeholder are excluded.

---

## Advanced Features

### Tables

```markdown
<!-- placeholder="TableArea" -->
| Left | Center | Right |
|:-----|:------:|------:|
| Alpha | Bravo | 1 |
| Charlie | Delta | 2 |
```

- The table is created at the position and size of the placeholder shape.
- Column widths are proportional to the dash count in the separator row.
- Alignment: `:---` left, `:--:` center, `---:` right.
- Tables without a placeholder comment are ignored.

---

### Images

```markdown
<!-- placeholder="ImageArea" -->
![Caption text](image.jpg)
```

- The image is fitted inside the placeholder using contain mode (aspect ratio preserved).
- If alt text is provided and a shape named `ImageArea_caption` exists on the slide, the caption is written there.
- Images without a placeholder comment are ignored.

---

### Table of Contents Slide

Add `toc: true` to the YAML front matter to generate a TOC slide automatically as the last slide.

```markdown
---
title: My Presentation
toc: true
---
```

- `##` (Section Header) and `###` (Title and Content) headings become TOC entries.
- Each entry is a hyperlink that jumps to the corresponding slide.

---

## Troubleshooting

### Log Levels

| Level | Default | With `-v` | Description |
|---|---|---|---|
| `DEBUG` | ❌ | ✅ | Shape info, placeholder resolution detail |
| `INFO` | ❌ | ✅ | Progress messages |
| `WARNING` | ✅ | ✅ | Non-fatal issues (layout not found, Pillow not installed, etc.) |
| `ERROR` | ✅ | ✅ | Fatal errors (file not found, save failed, etc.) |

---

### Common Issues

**Template not found**

```
File not found.
```

Provide the correct path with `--template`, or place `template.pptx` in the same directory as the script.

**Layout not found**

```
WARNING: Slide 2: layout 'MyLayout' not found. Falling back to auto.
```

Check that the layout name in your Markdown matches the slide layout name in the template exactly (case-sensitive). The script continues with an automatic fallback rather than crashing.

**Placeholder not found**

```
[page 2] placeholder "Content" NOT FOUND -> rescuing content
```

Open the PowerPoint template, check the shape name in the Selection Pane, and make sure it matches the name in your `<!-- placeholder="..." -->` comment exactly.

**Image not inserted**

```
WARNING: Image not found: logo.png
```

Place the image in the same directory as the Markdown file, or use an absolute path. Run with `-v` to see all paths that were searched.

**Formatting tags not applied**

```
WARNING: Slide 3: Unclosed tag in 'some <b>text'. Skipping formatting.
```

Verify that every opening `<b>`, `<i>`, or `<u>` tag has a matching closing tag in the correct order.

---

### Debug Workflow

```bash
# 1. Run with verbose flag
python md-pptx-injector.py input.md output.pptx -v

# 2. Check shapes on each slide
# [shapes on slide]
#   - #0: name='Title 1', is_placeholder=True, has_text_frame=True
#   - #1: name='Content 2', is_placeholder=True, has_text_frame=True

# 3. Verify placeholder resolution
# [page 2] placeholder "Content" found: actual_name='Content 2'

# 4. Verify image paths
# [image] inserted '/full/path/to/image.jpg' alt='Caption'
```

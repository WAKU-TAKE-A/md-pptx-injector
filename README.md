# md-pptx-injector.py

Markdown to PowerPoint converter with template support, custom layouts, and advanced formatting.

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

# With verbose output for debugging
python md-pptx-injector.py input.md output.pptx --template template.pptx -v
```

---

## Features

### Core Functionality
- ✅ Convert Markdown to PowerPoint presentations
- ✅ Template-based slide generation
- ✅ Custom layout support via HTML comments
- ✅ Custom placeholder targeting for precise content placement
- ✅ **Bold**, *Italic*, and ***Bold+Italic*** text formatting (inline and whole-line)
- ✅ Bullet points with multi-level indentation (up to 3 levels)
- ✅ Markdown tables → PowerPoint tables
- ✅ Images with aspect ratio preservation and captions
- ✅ YAML front matter for title slides
- ✅ PyInstaller packaging support

### Code Quality
- ✅ Modern Python 3.9+ type hints
- ✅ Comprehensive logging with adjustable verbosity
- ✅ Robust error handling with detailed diagnostics
- ✅ Type checking ready (mypy compatible)
- ✅ Clean, modular code structure

---

## Installation

### Requirements
- Python 3.9 or later
- Dependencies:
  ```bash
  pip install python-pptx Pillow
  ```

---

## Usage

### Command Line

```bash
python md-pptx-injector.py <src_md> <dst_pptx> [OPTIONS]
```

**Arguments:**
- `src_md` - Source Markdown file
- `dst_pptx` - Destination PowerPoint file

**Options:**
- `--template PATH` - Template PPTX file (default: `template.pptx`)
- `-v, --verbose` - Enable verbose logging for debugging

**Examples:**
```bash
# Use default template.pptx
python md-pptx-injector.py slides.md presentation.pptx

# Specify custom template
python md-pptx-injector.py slides.md presentation.pptx --template custom.pptx

# Debug mode with detailed logging
python md-pptx-injector.py slides.md presentation.pptx -v
```

---

## Path Resolution

### Application Base Directory
The "app base directory" is determined by:
- **Script execution**: Directory containing `md-pptx-injector.py`
- **EXE execution**: Directory containing `md-pptx-injector.exe` (PyInstaller)

### File Path Resolution (src_md, dst_pptx, --template)

Paths are resolved in this order:

1. **Absolute paths**: Used as-is (highest priority)
   ```bash
   python md-pptx-injector.py /home/user/slides.md /tmp/out.pptx
   ```

2. **Paths starting with `./` or `.\`**: Relative to current working directory
   ```bash
   python md-pptx-injector.py ./slides.md ./output.pptx
   ```

3. **Other relative paths**: Relative to app base directory
   ```bash
   python md-pptx-injector.py slides.md output.pptx
   # Looks for: <app_dir>/slides.md, <app_dir>/output.pptx
   ```

### Image Path Resolution

Images referenced in Markdown are resolved in this order:

1. **Absolute paths**: Used as-is
2. **Relative paths**: Search in order:
   - Markdown file directory (highest priority)
   - App base directory
   - Current working directory

---

## Markdown Syntax

### Page Breaks (Slide Boundaries)

Use `---` on a line by itself to create a new slide:

```markdown
# First Slide

Content for first slide

---

# Second Slide

Content for second slide
```

**Exception**: YAML front matter at the beginning of a page is not treated as a page break.

### Text Formatting

#### Bold
```markdown
**Bold text**
__Bold text__

**Whole line bold**
```

#### Italic
```markdown
*Italic text*
_Italic text_

*Whole line italic*
```

#### Bold + Italic
```markdown
***Bold and italic text***
___Bold and italic text___

***Whole line bold and italic***
```

#### Inline Formatting
```markdown
This is **bold**, *italic*, and ***bold+italic*** in one line.

You can mix **multiple** bold and *multiple* italic ***combinations***.
```

### Headings

#### Slide Layout Headings
```markdown
# Title Slide        → Uses "Title Slide" layout
## Section Header    → Uses "Section Header" layout
### Content Slide    → Uses "Title and Content" layout
```

#### In-Slide Headings (with level inheritance)
```markdown
#### Level 0 Heading
Content at level 0

##### Level 1 Heading
Content at level 1
- Bullet at level 1
  - Bullet at level 2
```

**Important**: Heading levels (#### and #####) affect subsequent content:
- `####` sets base level to 0
- `#####` sets base level to 1
- Regular paragraphs inherit the base level
- Bullet indentation adds to the base level (max level: 2)

### Bullet Points

```markdown
- Level 0 bullet
  - Level 1 bullet (2 spaces indent)
    - Level 2 bullet (4 spaces indent)

* Alternative marker
+ Another alternative
```

**Rules**:
- Indent: 2 spaces per level
- Maximum: 3 levels (0, 1, 2)
- Over-indentation is clipped to level 2

### Numbered Lists

```markdown
1. First item
2. Second item
3. Third item
```

**Note**: Numbers are treated as literal text (auto-numbering is not used).

### Blank Lines

Blank lines in Markdown create empty paragraphs in PowerPoint (vertical spacing).

---

## Layouts and Placeholders

### Layout Selection Priority

1. **Explicit layout comment** (highest priority):
   ```markdown
   <!-- layout="Custom Layout Name" -->
   ```

2. **Heading level**:
   - `#` → `Title Slide`
   - `##` → `Section Header`
   - `###` → `Title and Content`

3. **YAML front matter** (no heading):
   - Treated as `Title Slide`

### Title Slide

**Pattern A: YAML Front Matter**
```markdown
---
title: Presentation Title
subtitle: Subtitle Text
author: Author Name
---
```

**Pattern B: Markdown Headings** (overrides Pattern A)
```markdown
# Presentation Title
subtitle: Subtitle Text
author: Author Name
```

**Rendering**:
- Title → Title placeholder
- Subtitle + Author → Subtitle placeholder (combined, line break between)

### Custom Placeholders

Target specific placeholders by name:

```markdown
<!-- placeholder="Content" -->
This goes into the "Content" placeholder
(until blank line)

<!-- placeholder="SideNote" -->
This goes into the "SideNote" placeholder

<!-- placeholder="Content" -->
This is appended to "Content" placeholder
(with blank line separator)
```

**Rules**:
- Content is captured until a blank line
- Same placeholder name → append with blank line separator
- Empty block (immediate blank line) → adds vertical spacing
- Placeholders not found → ignored (no error)

### Rescue Content

Content not assigned to any placeholder is "rescued" into the body placeholder:

```markdown
<!-- layout="Two Content" -->

<!-- placeholder="LeftBox" -->
Goes to LeftBox

This unassigned text is rescued into the body placeholder
```

**Rescue Rules**:
- Only for non-Title Slide layouts
- Skipped if rescue content is only blank lines
- Avoids overwriting explicitly targeted placeholders

---

## Advanced Features

### Tables

**Markdown**:
```markdown
<!-- placeholder="TablePlaceholder" -->
| Left Heading | Centre Heading | Right Heading |
|:-------------|:--------------:|--------------:|
| Alpha        | Bravo          | 1             |
| Charlie      | Delta          | 2             |
```

**PowerPoint Result**:
- Creates table at placeholder position/size
- Column width: Based on dash count in separator row
- Alignment: `:---` (left), `:--:` (center), `---:` (right)

**Requirements**:
- Must be in a placeholder block
- Unprefixed tables are ignored

### Images

**Markdown**:
```markdown
<!-- placeholder="ImagePlaceholder" -->
![Caption text here](image.jpg)
```

**PowerPoint Result**:
- Image fitted to placeholder (contain mode)
- Aspect ratio preserved
- Caption → `ImagePlaceholder_caption` if exists

**Image Path Resolution**:
1. Markdown directory
2. App base directory
3. Current working directory

**Requirements**:
- Must be in a placeholder block
- Unprefixed images are ignored

---

## Recent Improvements

### Type Hints and Code Quality
- Migrated to Python 3.9+ modern type hints (`list`, `dict`, `| None`)
- Added comprehensive type checking support (mypy.ini included)
- Replaced generic `Exception` with specific exception types
- Refactored long functions into single-responsibility modules

### Logging System
- Replaced custom logging with Python standard `logging` module
- **Default**: `WARNING` level (errors and warnings only)
- **Verbose mode** (`-v`): `DEBUG` level (all diagnostic info)

```bash
# Normal operation (quiet)
python md-pptx-injector.py input.md output.pptx

# Debug mode (verbose)
python md-pptx-injector.py input.md output.pptx -v
```

### Error Messages
Enhanced error reporting with diagnostic information:
```
ERROR: Source markdown file not found: /path/to/file.md
  Current working directory: /current/dir
  App directory: /app/dir
  Absolute path: /path/to/file.md
```

### Image Path Resolution
Improved multi-path fallback with detailed logging:
```
Image not found: logo.png
  Searched paths:
    - /markdown/dir/logo.png
    - /app/dir/logo.png
    - /cwd/logo.png
```

### Table Parsing
Enhanced validation with better error messages:
- Column count verification
- Separator row validation
- Detailed debug logging for malformed tables

### Bold/Italic Support
Full inline and whole-line formatting:
- `**bold**`, `*italic*`, `***bold+italic***`
- `__bold__`, `_italic_`, `___bold+italic___`
- Mixed formatting in single line
- Negative lookahead/lookbehind to avoid pattern conflicts

---

## Troubleshooting

### Logging Levels

| Level | Default | Verbose | Usage |
|-------|---------|---------|-------|
| `DEBUG` | ❌ | ✅ | Detailed trace (shape info, placeholder resolution) |
| `INFO` | ❌ | ✅ | Progress messages |
| `WARNING` | ✅ | ✅ | Non-critical issues (missing placeholders, Pillow not installed) |
| `ERROR` | ✅ | ✅ | Critical errors (file not found, save failed) |

### Common Issues

**1. Template not found**
```
ERROR: Template pptx file not found: template.pptx
```
**Solution**: Provide template with `--template` or place `template.pptx` in app directory

**2. Layout not found**
```
ValueError: Layout not found in template: "CustomLayout"
```
**Solution**: Check layout name matches exactly (case-sensitive) in PowerPoint template

**3. Image not loaded**
```
WARNING: Image not found: image.jpg
```
**Solution**: Check image path relative to Markdown file or use absolute path

**4. Placeholder not found**
```
[page 2] placeholder "Content" NOT FOUND
```
**Solution**: Verify placeholder name in PowerPoint selection pane matches exactly

**5. Bold+Italic not working**
```markdown
# Problem
***text***  → Only bold applied

# Solution (after fix)
***text***  → Bold+italic applied correctly
```
**Solution**: Use latest version with fixed regex patterns

### Debug Workflow

1. **Run with verbose flag**:
   ```bash
   python md-pptx-injector.py input.md output.pptx -v
   ```

2. **Check slide shapes**:
   ```
   [page 1] created slide: slide_layout.name='Title Slide'
   [shapes on slide]
     - #0: name='Title 1', is_placeholder=True, has_text_frame=True
     - #1: name='Subtitle 2', is_placeholder=True, has_text_frame=True
   ```

3. **Verify placeholder resolution**:
   ```
   [page 2] placeholder "Content" found: actual_name='Content Placeholder 2'
   ```

4. **Check image paths**:
   ```
   Image found: /full/path/to/image.jpg
   ```


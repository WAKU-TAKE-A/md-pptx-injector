#!/usr/bin/env python3
# conv-md-to-pptx.py
#
# Usage:
#   python conv-md-to-pptx.py src.md dst.pptx --template ref.pptx --verbose
#
# Path rules:
# - Absolute paths: used as-is (highest priority)
# - Paths starting with "./" or ".\": resolved relative to current working directory (cwd)
# - Other relative paths: resolved relative to the app directory
#   - script run: directory containing this .py
#   - exe run (PyInstaller etc): directory containing the .exe
#
# Dependencies:
#   pip install python-pptx Pillow

from __future__ import annotations

import argparse
import logging
import re
import sys
from dataclasses import dataclass
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement

try:
    from PIL import Image as PILImage
except ImportError:  # pragma: no cover
    PILImage = None

# Setup logger
logger = logging.getLogger(__name__)


# -------------------------
# Logging setup
# -------------------------
def setup_logging(verbose: bool) -> None:
    """Configure logging based on verbosity level."""
    level = logging.DEBUG if verbose else logging.WARNING
    logging.basicConfig(
        level=level,
        format='%(message)s',
        stream=sys.stdout
    )


# -------------------------
# App dir / path resolution
# -------------------------
def get_app_dir() -> Path:
    """Get application directory (supports PyInstaller packaging)."""
    # When packaged as exe (PyInstaller, etc.)
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    # When running as normal .py script
    return Path(__file__).resolve().parent


def resolve_path(base_dir: Path, p: str) -> Path:
    """
    Resolve path with priority:
    - Absolute path: as-is
    - Starts with ./ or .\: relative to cwd
    - Otherwise: relative to base_dir (app dir)
    """
    pp = Path(p)

    if pp.is_absolute():
        return pp

    if p.startswith("./") or p.startswith(".\\"):
        return (Path.cwd() / pp).resolve()

    return (base_dir / pp).resolve()


# -------------------------
# Regex patterns
# -------------------------
RE_LAYOUT = re.compile(r'<!--\s*layout\s*=\s*"([^"]+)"\s*--\s*>')
RE_PLACEHOLDER = re.compile(r'<!--\s*placeholder\s*=\s*"([^"]+)"\s*--\s*>')
RE_HEADING = re.compile(r"^(#{1,6})\s+(.*)$")
RE_KV = re.compile(r"^\s*(title|subtitle|author)\s*:\s*(.*)\s*$", re.IGNORECASE)
RE_BULLET = re.compile(r"^(\s*)([-*+])\s+(.*)$")
RE_IMAGE = re.compile(r"!\[(.*?)\]\((.*?)\)")
RE_STRONG_LINE1 = re.compile(r"^\s*\*\*(?!\*)(.+?)(?<!\*)\*\*\s*$")
RE_STRONG_LINE2 = re.compile(r"^\s*__(?!_)(.+?)(?<!_)__\s*$")

# Inline formatting patterns (for bold and italic)
RE_BOLD_ITALIC = re.compile(r'\*\*\*(.+?)\*\*\*|___(.+?)___')
RE_BOLD = re.compile(r'\*\*(.+?)\*\*|__(.+?)__')
RE_ITALIC = re.compile(r'\*(.+?)\*|_(.+?)_')


# -------------------------
# Logging helpers
# -------------------------
def dump_slide_shapes(slide, verbose: bool) -> None:
    """Debug: dump all shapes on a slide."""
    if not verbose:
        return
    logger.debug("  [shapes on slide]")
    for i, shp in enumerate(slide.shapes):
        name = getattr(shp, "name", "")
        has_tf = getattr(shp, "has_text_frame", False)
        is_ph = getattr(shp, "is_placeholder", False)

        ph_idx = None
        ph_type = None
        if is_ph:
            try:
                ph_idx = shp.placeholder_format.idx
                ph_type = shp.placeholder_format.type
            except (AttributeError, KeyError):
                pass

        st = getattr(shp, "shape_type", None)
        logger.debug(
            f"    - #{i}: name={name!r}, shape_type={st}, "
            f"is_placeholder={is_ph}, ph_idx={ph_idx}, ph_type={ph_type}, has_text_frame={has_tf}"
        )


# -------------------------
# PPTX helpers
# -------------------------
def _ppr(paragraph):
    """Get or add paragraph properties element."""
    return paragraph._p.get_or_add_pPr()


def set_bullet_none(paragraph) -> None:
    """Disable bullets for this paragraph (a:buNone)."""
    pPr = _ppr(paragraph)

    for tag in (
        "a:buNone",
        "a:buChar",
        "a:buAutoNum",
        "a:buBlip",
        "a:buClr",
        "a:buSzPct",
        "a:buSzPts",
        "a:buFont",
    ):
        for el in pPr.findall(qn(tag)):
            pPr.remove(el)

    buNone = OxmlElement("a:buNone")
    pPr.insert(0, buNone)


def clear_bullet_override(paragraph) -> None:
    """Remove explicit a:buNone so placeholder defaults can apply."""
    pPr = _ppr(paragraph)
    for el in pPr.findall(qn("a:buNone")):
        pPr.remove(el)


def bring_to_front(shape) -> None:
    """Ensure the shape is last in spTree (topmost)."""
    try:
        sp = shape._element
        parent = sp.getparent()
        parent.remove(sp)
        parent.append(sp)
    except (AttributeError, ValueError):
        pass


def find_layout_by_name(prs: Presentation, layout_name: str):
    """Find slide layout by name in presentation."""
    for m in prs.slide_masters:
        for layout in m.slide_layouts:
            if layout.name == layout_name:
                return layout
    raise ValueError(f'Layout not found in template: "{layout_name}"')


def find_title_shape(slide):
    """Find the title shape on a slide."""
    try:
        if slide.shapes.title is not None:
            return slide.shapes.title
    except (AttributeError, KeyError):
        pass

    for shp in getattr(slide, "placeholders", []):
        if not hasattr(shp, "placeholder_format"):
            continue
        try:
            if shp.placeholder_format.type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE):
                return shp
        except (AttributeError, KeyError):
            pass
    return None


def find_body_text_shape_excluding(slide, exclude_ids: set[int]):
    """
    Find a body-like placeholder, excluding shapes (by id()).
    Prevents "rescue" from overwriting explicitly-targeted placeholders.
    """
    def ph(name):
        return getattr(PP_PLACEHOLDER, name, None)

    skip = {
        ph("TITLE"),
        ph("CENTER_TITLE"),
        ph("SUBTITLE"),
        ph("DATE"),
        ph("FOOTER"),
        ph("HEADER"),
        ph("SLIDE_NUMBER"),
    }
    skip = {x for x in skip if x is not None}

    for shp in getattr(slide, "placeholders", []):
        if id(shp) in exclude_ids:
            continue
        if not getattr(shp, "has_text_frame", False):
            continue
        try:
            if shp.placeholder_format.type in skip:
                continue
        except (AttributeError, KeyError):
            pass
        return shp

    title = find_title_shape(slide)
    for shp in slide.shapes:
        if shp is title:
            continue
        if id(shp) in exclude_ids:
            continue
        if getattr(shp, "has_text_frame", False):
            return shp

    return None


def find_shape_by_name(slide, shape_name: str):
    """
    Find shape by name. Priority:
    1) slide.shapes name match
    2) layout placeholder name -> idx -> slide placeholder idx mapping
    """
    for shp in slide.shapes:
        if getattr(shp, "name", None) == shape_name:
            return shp

    try:
        idx = None
        for lph in slide.slide_layout.placeholders:
            if getattr(lph, "name", None) == shape_name:
                idx = lph.placeholder_format.idx
                break
        if idx is None:
            return None

        for sph in slide.placeholders:
            if sph.placeholder_format.idx == idx:
                return sph
    except (AttributeError, KeyError):
        pass

    return None


def set_text_lines(shape, lines: list[str]) -> None:
    """Set multi-line text into a text-frame shape without forcing font settings."""
    if shape is None or not getattr(shape, "has_text_frame", False):
        return
    tf = shape.text_frame
    tf.clear()

    if not lines:
        return

    p0 = tf.paragraphs[0]
    p0.text = lines[0]
    set_bullet_none(p0)

    for s in lines[1:]:
        p = tf.add_paragraph()
        p.text = s
        set_bullet_none(p)


# -------------------------
# Markdown parsing helpers
# -------------------------
def split_pages(md_text: str) -> list[list[str]]:
    """
    Split by a line that is exactly '---', except when it's a YAML front matter
    start/end at the beginning of a page.
    """
    lines = md_text.splitlines()
    pages: list[list[str]] = []
    cur: list[str] = []

    in_front_matter = False
    seen_nonblank_in_page = False
    front_matter_started = False

    def flush():
        nonlocal cur, in_front_matter, seen_nonblank_in_page, front_matter_started
        pages.append(cur)
        cur = []
        in_front_matter = False
        seen_nonblank_in_page = False
        front_matter_started = False

    for line in lines:
        is_delim = (line.strip() == "---" and line.strip() == line)

        if not seen_nonblank_in_page and line.strip() != "":
            seen_nonblank_in_page = True
            if is_delim:
                in_front_matter = True
                front_matter_started = True
                cur.append(line)
                continue

        if in_front_matter:
            cur.append(line)
            if is_delim and front_matter_started and len(cur) > 1:
                in_front_matter = False
            continue

        if is_delim:
            flush()
        else:
            cur.append(line)

    if cur or not pages:
        pages.append(cur)
    return pages


def extract_front_matter(page_lines: list[str]) -> tuple[dict[str, str], list[str]]:
    """Extract YAML front matter from page lines."""
    i = 0
    while i < len(page_lines) and page_lines[i].strip() == "":
        i += 1
    if i >= len(page_lines) or page_lines[i].strip() != "---":
        return {}, page_lines

    j = i + 1
    while j < len(page_lines):
        if page_lines[j].strip() == "---":
            yaml_lines = page_lines[i + 1 : j]
            rest = page_lines[j + 1 :]
            return parse_simple_yaml(yaml_lines), rest
        j += 1

    return {}, page_lines


def parse_simple_yaml(yaml_lines: list[str]) -> dict[str, str]:
    """Parse simple YAML key: value pairs."""
    data: dict[str, str] = {}
    for line in yaml_lines:
        m = RE_KV.match(line)
        if not m:
            continue
        key = m.group(1).lower()
        val = m.group(2).strip()
        data[key] = val
    return data


def find_first_heading(lines: list[str]) -> tuple[int | None, str | None, int, str | None]:
    """Find first markdown heading in lines."""
    for idx, line in enumerate(lines):
        s = line.strip()
        m = RE_HEADING.match(s)
        if not m:
            continue
        level = len(m.group(1))
        text = m.group(2).strip()
        return level, text, idx, s
    return None, None, -1, None


def extract_layout_override(lines: list[str]) -> str | None:
    """Extract layout override from HTML comment."""
    for ln in lines:
        m = RE_LAYOUT.search(ln)
        if m:
            return m.group(1).strip()
    return None


def strip_layout_comment_lines(lines: list[str]) -> list[str]:
    """Remove layout comment lines from content."""
    return [ln for ln in lines if not RE_LAYOUT.search(ln)]


def has_nonblank_text(lines: list[str]) -> bool:
    """Check if lines contain any non-blank text."""
    return any(ln.strip() != "" for ln in lines)


@dataclass
class PlaceholderBlock:
    """Represents a placeholder block in markdown."""
    name: str
    lines: list[str]  # captured until blank line


def parse_placeholder_blocks(lines: list[str]) -> tuple[dict[str, list[PlaceholderBlock]], list[str]]:
    """
    Parse placeholder blocks:
      <!-- placeholder="X" -->
      ... (until blank line)

    - The blank line that terminates the block is consumed and NOT added to rescue.
    """
    blocks: dict[str, list[PlaceholderBlock]] = {}
    rescue: list[str] = []

    i = 0
    while i < len(lines):
        ln = lines[i]

        phm = RE_PLACEHOLDER.search(ln)
        if phm:
            name = phm.group(1).strip()
            i += 1
            captured: list[str] = []
            while i < len(lines) and lines[i].strip() != "":
                captured.append(lines[i])
                i += 1

            blocks.setdefault(name, []).append(PlaceholderBlock(name=name, lines=captured))

            # consume ONE terminating blank line (do not put into rescue)
            if i < len(lines) and lines[i].strip() == "":
                i += 1
            continue

        if RE_LAYOUT.search(ln):
            i += 1
            continue

        if RE_PLACEHOLDER.search(ln):
            i += 1
            continue

        rescue.append(ln)
        i += 1

    return blocks, rescue


# -------------------------
# Inline formatting (bold/italic)
# -------------------------
@dataclass
class TextRun:
    """Represents a text run with formatting."""
    text: str
    bold: bool = False
    italic: bool = False


def parse_inline_formatting(text: str) -> list[TextRun]:
    """
    Parse inline markdown formatting (bold, italic, bold+italic).
    
    Patterns:
    - ***text*** or ___text___ -> bold + italic
    - **text** or __text__ -> bold
    - *text* or _text_ -> italic
    """
    runs: list[TextRun] = []
    pos = 0
    
    # Build a combined pattern that tries bold+italic first, then bold, then italic
    # Use negative lookahead to avoid matching *** as ** + *
    combined_pattern = re.compile(
        r'\*\*\*(?P<bi1>.+?)\*\*\*|'  # ***text***
        r'___(?P<bi2>.+?)___|'         # ___text___
        r'\*\*(?P<b1>.+?)\*\*|'        # **text**
        r'__(?P<b2>.+?)__|'            # __text__
        r'\*(?P<i1>.+?)\*|'            # *text*
        r'_(?P<i2>.+?)_'               # _text_
    )
    
    for match in combined_pattern.finditer(text):
        # Add text before match as plain text
        if pos < match.start():
            runs.append(TextRun(text=text[pos:match.start()]))
        
        # Determine which group matched
        if match.group('bi1') or match.group('bi2'):
            # Bold + Italic
            content = match.group('bi1') or match.group('bi2')
            runs.append(TextRun(text=content, bold=True, italic=True))
        elif match.group('b1') or match.group('b2'):
            # Bold only
            content = match.group('b1') or match.group('b2')
            runs.append(TextRun(text=content, bold=True))
        elif match.group('i1') or match.group('i2'):
            # Italic only
            content = match.group('i1') or match.group('i2')
            runs.append(TextRun(text=content, italic=True))
        
        pos = match.end()
    
    # Add remaining text
    if pos < len(text):
        runs.append(TextRun(text=text[pos:]))
    
    # If no formatting found, return single run with original text
    if not runs:
        runs.append(TextRun(text=text))
    
    return runs


# -------------------------
# Markdown -> paragraph specs
# -------------------------
@dataclass
class ParaSpec:
    """Paragraph specification with formatting."""
    text: str
    bullet: bool = False
    level: int = 0     # used for both bullets and template-level styling
    bold: bool = False
    empty: bool = False
    italic: bool = False
    runs: list[TextRun] | None = None  # For inline formatting


def skip_tables_and_images(lines: list[str]) -> list[str]:
    """
    Remove markdown tables/images from normal body flow.
    Tables/images are only processed when custom placeholder specified.
    """
    out: list[str] = []
    i = 0
    while i < len(lines):
        ln = lines[i]

        if RE_IMAGE.search(ln.strip()):
            i += 1
            continue

        if ln.strip().startswith("|") and i + 1 < len(lines):
            sep = lines[i + 1].strip()
            if sep.startswith("|") and "-" in sep:
                i += 2
                while i < len(lines) and lines[i].strip() != "":
                    i += 1
                continue

        out.append(ln)
        i += 1
    return out


def build_paragraphs_from_lines(lines: list[str]) -> list[ParaSpec]:
    """
    Stateful parsing:
    - Tracks current base_level set by #### / #####
    - Subsequent paragraphs inherit base_level
    - Bullet items level = min(base_level + indent_level, 2)
      - indent_level: 1 per 2 spaces
      - max levels total: 0..2
    """
    paras: list[ParaSpec] = []
    lines = skip_tables_and_images(lines)

    base_level = 0

    for raw in lines:
        line = raw.rstrip("\n")

        if line.strip() == "":
            paras.append(ParaSpec(text="", empty=True, bullet=False, level=base_level))
            continue

        # #### / ##### set base_level and emit a bold heading paragraph
        hm = RE_HEADING.match(line.strip())
        if hm and len(hm.group(1)) in (4, 5):
            text = hm.group(2).strip()
            base_level = 0 if len(hm.group(1)) == 4 else 1
            paras.append(ParaSpec(text=text, bullet=False, bold=True, level=base_level))
            continue

        # whole-line bold
        sm = RE_STRONG_LINE1.match(line) or RE_STRONG_LINE2.match(line)
        if sm:
            paras.append(ParaSpec(text=sm.group(1).strip(), bullet=False, bold=True, level=base_level))
            continue

        # bullets
        bm = RE_BULLET.match(line)
        if bm:
            indent = len(bm.group(1).replace("\t", "  "))
            indent_level = indent // 2
            level = min(base_level + indent_level, 2)
            text = bm.group(3).strip()
            
            # Parse inline formatting for bullet text
            runs = parse_inline_formatting(text)
            paras.append(ParaSpec(text=text, bullet=True, level=level, runs=runs))
            continue

        # normal paragraph inherits base_level
        # Parse inline formatting
        text = line.strip()
        runs = parse_inline_formatting(text)
        paras.append(ParaSpec(text=text, bullet=False, level=base_level, runs=runs))

    return paras


def write_paragraphs_to_shape(shape, paras: list[ParaSpec], append: bool, blank_before_append: bool) -> None:
    """Write paragraphs to a text shape with formatting support."""
    if shape is None or not getattr(shape, "has_text_frame", False):
        return

    tf = shape.text_frame

    if not append:
        tf.clear()
        first = tf.paragraphs[0]
    else:
        first = None

    def _new_paragraph():
        return tf.add_paragraph()

    if append and blank_before_append:
        p = _new_paragraph()
        p.text = ""
        set_bullet_none(p)
        p.level = 0

    if not paras:
        return

    for idx, ps in enumerate(paras):
        if first is not None and idx == 0:
            p = first
        else:
            p = _new_paragraph()

        # Handle inline formatting if present
        if ps.runs and len(ps.runs) > 0:
            # Clear existing text
            p.text = ""
            
            # Add formatted runs
            for run_spec in ps.runs:
                run = p.add_run()
                run.text = run_spec.text
                
                try:
                    if run_spec.bold:
                        run.font.bold = True
                    if run_spec.italic:
                        run.font.italic = True
                except (AttributeError, KeyError):
                    pass
        else:
            # Simple text without inline formatting
            p.text = ps.text
            
            # Apply paragraph-level bold
            if ps.bold:
                try:
                    p.font.bold = True
                except (AttributeError, KeyError):
                    if p.runs:
                        p.runs[0].font.bold = True
            
            # Apply paragraph-level italic
            if ps.italic:
                try:
                    p.font.italic = True
                except (AttributeError, KeyError):
                    if p.runs:
                        p.runs[0].font.italic = True

        # Set bullet/level
        if ps.bullet:
            clear_bullet_override(p)
            p.level = ps.level
        else:
            set_bullet_none(p)
            p.level = ps.level  # template-driven size differences


# -------------------------
# Table handling
# -------------------------
def split_pipe_row(row: str) -> list[str]:
    """Split a markdown table row by pipes."""
    s = row.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        s = s[:-1]
    return [c.strip() for c in s.split("|")]


def parse_markdown_table(lines: list[str]) -> tuple[list[str], list[list[str]], list[int], list[PP_ALIGN]] | None:
    """
    Parse markdown table with validation.
    
    Returns:
        Tuple of (headers, data_rows, width_units, alignments) or None if invalid
    """
    tbl = [ln.strip() for ln in lines if ln.strip() != ""]
    if len(tbl) < 2:
        logger.debug("Table parsing: too short (needs at least header + separator)")
        return None
    
    if not (tbl[0].startswith("|") and tbl[1].startswith("|")):
        logger.debug("Table parsing: doesn't start with |")
        return None
    
    if "-" not in tbl[1]:
        logger.debug("Table parsing: separator row missing dashes")
        return None

    headers = split_pipe_row(tbl[0])
    sep_cells = split_pipe_row(tbl[1])
    
    if len(sep_cells) != len(headers):
        logger.warning(
            f"Table parsing: column count mismatch - "
            f"headers={len(headers)}, separator={len(sep_cells)}"
        )
        return None

    width_units: list[int] = []
    aligns: list[PP_ALIGN] = []
    for cell in sep_cells:
        dash_count = cell.replace(":", "").count("-")
        if dash_count == 0:
            logger.debug(f"Table parsing: invalid separator cell: {cell}")
            return None
        width_units.append(max(dash_count, 1))

        c = cell.strip()
        left = c.startswith(":")
        right = c.endswith(":")
        if left and right:
            aligns.append(PP_ALIGN.CENTER)
        elif right:
            aligns.append(PP_ALIGN.RIGHT)
        else:
            aligns.append(PP_ALIGN.LEFT)

    data_rows: list[list[str]] = []
    for ln in tbl[2:]:
        if not ln.startswith("|"):
            continue
        r = split_pipe_row(ln)
        if len(r) < len(headers):
            r += [""] * (len(headers) - len(r))
        elif len(r) > len(headers):
            r = r[: len(headers)]
        data_rows.append(r)

    return headers, data_rows, width_units, aligns


def insert_table_at_shape(slide, shape, table_lines: list[str]) -> None:
    """Insert a markdown table at the given shape's position."""
    parsed = parse_markdown_table(table_lines)
    if parsed is None:
        logger.warning("Failed to parse markdown table")
        return
    headers, rows, units, aligns = parsed
    if shape is None:
        return

    left, top, width, height = shape.left, shape.top, shape.width, shape.height
    n_rows = 1 + len(rows)
    n_cols = len(headers)

    tbl_shape = slide.shapes.add_table(n_rows, n_cols, left, top, width, height)
    bring_to_front(tbl_shape)
    tbl = tbl_shape.table

    total_units = sum(units) if sum(units) > 0 else n_cols
    used = 0
    for ci in range(n_cols):
        if ci == n_cols - 1:
            w = width - used
        else:
            w = int(width * units[ci] / total_units)
            used += w
        try:
            tbl.columns[ci].width = w
        except (AttributeError, IndexError):
            pass

    for ci, h in enumerate(headers):
        cell = tbl.cell(0, ci)
        cell.text = h
        try:
            cell.text_frame.paragraphs[0].alignment = aligns[ci]
        except (AttributeError, IndexError):
            pass

    for ri, row in enumerate(rows, start=1):
        for ci, val in enumerate(row):
            cell = tbl.cell(ri, ci)
            cell.text = val
            try:
                cell.text_frame.paragraphs[0].alignment = aligns[ci]
            except (AttributeError, IndexError):
                pass


# -------------------------
# Image handling
# -------------------------
def parse_single_image_line(lines: list[str]) -> tuple[str, str] | None:
    """Parse a single image markdown line."""
    tbl = [ln.strip() for ln in lines if ln.strip() != ""]
    if len(tbl) != 1:
        return None
    m = RE_IMAGE.search(tbl[0])
    if not m:
        return None
    alt = (m.group(1) or "").strip()
    path = (m.group(2) or "").strip()
    return alt, path


def add_picture_contain(slide, img_path: Path, box_left, box_top, box_w, box_h):
    """Add picture with contain fit (preserving aspect ratio)."""
    if PILImage is None:
        logger.warning("Pillow not installed; skipping image insertion.")
        return None

    try:
        with PILImage.open(img_path) as im:
            w_px, h_px = im.size
    except (IOError, OSError) as e:
        logger.error(f"Failed to open image: {img_path} ({e})")
        return None

    if w_px <= 0 or h_px <= 0:
        return None

    img_ar = w_px / h_px
    box_ar = box_w / box_h if box_h else img_ar

    if img_ar >= box_ar:
        new_w = box_w
        new_h = int(box_w / img_ar)
    else:
        new_h = box_h
        new_w = int(box_h * img_ar)

    left = box_left + int((box_w - new_w) / 2)
    top = box_top + int((box_h - new_h) / 2)

    pic = slide.shapes.add_picture(str(img_path), left, top, width=new_w, height=new_h)
    bring_to_front(pic)
    return pic


def resolve_image_path(
    img_rel: str, 
    md_dir: Path, 
    app_dir: Path
) -> Path | None:
    """
    Resolve image path with multiple fallback strategies.
    Returns None if image not found.
    
    Search order:
    1) Absolute path (if provided)
    2) Relative to markdown file directory
    3) Relative to app directory
    4) Relative to current working directory
    """
    p = Path(img_rel)
    
    search_paths: list[Path] = []
    
    if p.is_absolute():
        search_paths.append(p)
    else:
        # 1) Markdown directory (highest priority)
        search_paths.append((md_dir / p).resolve())
        # 2) App directory
        search_paths.append((app_dir / p).resolve())
        # 3) Current working directory
        search_paths.append((Path.cwd() / p).resolve())
    
    for path in search_paths:
        if path.exists():
            logger.debug(f"Image found: {path}")
            return path
    
    logger.warning(
        f"Image not found: {img_rel}\n"
        f"  Searched paths:\n" + 
        "\n".join(f"    - {p}" for p in search_paths)
    )
    return None


def insert_image_at_shape(
    slide,
    shape,
    img_line_block: list[str],
    placeholder_name: str,
    md_dir: Path,
    app_dir: Path,
    verbose: bool,
) -> None:
    """Insert image at shape position."""
    parsed = parse_single_image_line(img_line_block)
    if parsed is None:
        return
    alt, img_rel = parsed

    if shape is None:
        return

    img_path = resolve_image_path(img_rel, md_dir=md_dir, app_dir=app_dir)
    if img_path is None:
        return

    logger.debug(f'[image] placeholder="{placeholder_name}" path="{img_path}" contain-fit')
    add_picture_contain(slide, img_path, shape.left, shape.top, shape.width, shape.height)

    if alt.strip():
        cap_shape = find_shape_by_name(slide, f"{placeholder_name}_caption")
        if cap_shape is not None and getattr(cap_shape, "has_text_frame", False):
            set_text_lines(cap_shape, [alt.strip()])


# -------------------------
# Slide building - helper functions
# -------------------------
def parse_title_page_info(front: dict[str, str], body_lines: list[str]) -> dict[str, str]:
    """
    Extract title page info:
    - Pattern A: front matter (title/subtitle/author)
    - Pattern B: body (# title + subtitle:/author:), overrides A (later wins).
    """
    info = dict(front)
    for ln in body_lines:
        m = RE_HEADING.match(ln.strip())
        if m and len(m.group(1)) == 1:
            info["title"] = m.group(2).strip()
            continue
        km = RE_KV.match(ln)
        if km:
            key = km.group(1).lower()
            val = km.group(2).strip()
            if key in ("subtitle", "author", "title"):
                info[key] = val
    return info


def determine_layout_name(layout_override: str | None, heading_level: int | None, has_front: bool) -> str:
    """Determine slide layout name based on content."""
    if layout_override:
        return layout_override
    if heading_level == 1:
        return "Title Slide"
    if heading_level == 2:
        return "Section Header"
    if heading_level == 3:
        return "Title and Content"
    if has_front:
        return "Title Slide"
    return "Title and Content"


def remove_first_exact_line(lines: list[str], target: str | None) -> list[str]:
    """Remove first occurrence of exact line match."""
    if not target:
        return lines
    out = []
    removed = False
    for ln in lines:
        if not removed and ln.strip() == target:
            removed = True
            continue
        out.append(ln)
    return out


def build_title_slide(
    slide, 
    front: dict[str, str], 
    body_lines: list[str], 
    title_shape,
    verbose: bool
) -> None:
    """Build Title Slide layout."""
    info = parse_title_page_info(front, body_lines)

    title = info.get("title", "").strip()
    subtitle = info.get("subtitle", "").strip()
    author = info.get("author", "").strip()

    if title_shape is not None and title:
        set_text_lines(title_shape, [title])

    # Find subtitle placeholder
    sub_shape = None
    for shp in getattr(slide, "placeholders", []):
        if getattr(shp, "has_text_frame", False):
            try:
                if shp.placeholder_format.type == PP_PLACEHOLDER.SUBTITLE:
                    sub_shape = shp
                    break
            except (AttributeError, KeyError):
                pass

    if sub_shape is None:
        for shp in slide.shapes:
            if shp is title_shape:
                continue
            if getattr(shp, "has_text_frame", False):
                sub_shape = shp
                break

    lines = []
    if subtitle:
        lines.append(subtitle)
    if author:
        lines.append(author)
    if sub_shape is not None and lines:
        set_text_lines(sub_shape, lines)


def build_content_slide(
    slide,
    heading_text: str | None,
    heading_level: int | None,
    heading_raw: str | None,
    rescue_lines: list[str],
    title_shape
) -> list[str]:
    """
    Build content slide (non-title slide).
    Returns updated rescue_lines.
    """
    if title_shape is not None and heading_text and heading_level in (2, 3):
        set_text_lines(title_shape, [heading_text])
    return remove_first_exact_line(rescue_lines, heading_raw)


def process_placeholder_blocks(
    slide,
    blocks_by_name: dict[str, list[PlaceholderBlock]],
    md_dir: Path,
    app_dir: Path,
    state: dict,
    page_no: int,
    verbose: bool
) -> set[int]:
    """
    Process all placeholder blocks (text/table/image).
    Returns set of explicitly used shape IDs.
    """
    explicitly_used_shape_ids: set[int] = set()

    for ph_name, blocks in blocks_by_name.items():
        shp = find_shape_by_name(slide, ph_name)

        if shp is None:
            logger.debug(f'[page {page_no}] placeholder "{ph_name}" NOT FOUND')
            continue

        explicitly_used_shape_ids.add(id(shp))

        logger.debug(
            f'[page {page_no}] placeholder "{ph_name}" found: actual_name={getattr(shp, "name", None)!r}, '
            f'is_placeholder={getattr(shp, "is_placeholder", False)}, has_text_frame={getattr(shp, "has_text_frame", False)}'
        )

        for bi, blk in enumerate(blocks):
            # table
            if parse_markdown_table(blk.lines) is not None:
                logger.debug(f'  [table] block #{bi+1} -> "{ph_name}"')
                try:
                    insert_table_at_shape(slide, shp, blk.lines)
                except Exception as e:
                    logger.error(f'Table insertion failed for "{ph_name}": {e}')
                continue

            # image
            if parse_single_image_line(blk.lines) is not None:
                logger.debug(f'  [image] block #{bi+1} -> "{ph_name}"')
                try:
                    insert_image_at_shape(
                        slide,
                        shp,
                        blk.lines,
                        ph_name,
                        md_dir=md_dir,
                        app_dir=app_dir,
                        verbose=verbose,
                    )
                except Exception as e:
                    logger.error(f'Image insertion failed for "{ph_name}": {e}')
                continue

            # text
            if not getattr(shp, "has_text_frame", False):
                logger.debug(f'  [text] block #{bi+1} -> "{ph_name}" skipped (no text frame)')
                continue

            paras = build_paragraphs_from_lines(blk.lines)
            already = state["text_written"].get((id(slide), ph_name), False)
            logger.debug(f'  [text] block #{bi+1} -> "{ph_name}" paras={len(paras)} append={already}')
            write_paragraphs_to_shape(shp, paras, append=already, blank_before_append=already)
            state["text_written"][(id(slide), ph_name)] = True

    return explicitly_used_shape_ids


def process_rescue_content(
    slide,
    rescue_lines: list[str],
    explicitly_used_shape_ids: set[int],
    state: dict,
    page_no: int,
    verbose: bool
) -> None:
    """Process rescue content (unspecified text) into body placeholder."""
    # Skip if rescue contains only blank lines (avoid overwriting explicit placeholders with empty content)
    if not has_nonblank_text(rescue_lines):
        logger.debug(f"[page {page_no}] rescue skipped (no nonblank text)")
        return

    body_shape = find_body_text_shape_excluding(slide, explicitly_used_shape_ids)
    if body_shape is None:
        logger.debug(f"[page {page_no}] body placeholder NOT FOUND (rescue skipped)")
        return

    rescue_paras = build_paragraphs_from_lines(rescue_lines)
    if rescue_paras:
        key = (id(slide), "__BODY__")
        already = state["text_written"].get(key, False)
        logger.debug(f"[page {page_no}] rescue -> body paras={len(rescue_paras)} append={already}")
        write_paragraphs_to_shape(body_shape, rescue_paras, append=already, blank_before_append=already)
        state["text_written"][key] = True


# -------------------------
# Slide build (per page) - main function
# -------------------------
def build_slide_from_page(
    prs: Presentation,
    page_lines: list[str],
    md_dir: Path,
    app_dir: Path,
    state: dict,
    page_no: int = 0,
) -> None:
    """
    Build a single slide from markdown page content.
    
    This is the main orchestration function that:
    1. Parses front matter and determines layout
    2. Creates the slide
    3. Builds title or content slide
    4. Processes placeholder blocks
    5. Processes rescue content
    """
    verbose = bool(state.get("verbose", False))

    front, body0 = extract_front_matter(page_lines)
    body0_wo_layout = strip_layout_comment_lines(body0)
    layout_override = extract_layout_override(page_lines)

    heading_level, heading_text, heading_idx, heading_raw = find_first_heading(body0_wo_layout)
    layout_name = determine_layout_name(layout_override, heading_level, bool(front))

    logger.debug(
        f"[page {page_no}] heading=({heading_level}) {heading_text!r}, layout_override={layout_override!r} => layout={layout_name!r}"
    )

    layout = find_layout_by_name(prs, layout_name)
    slide = prs.slides.add_slide(layout)

    logger.debug(f"[page {page_no}] created slide: slide_layout.name={slide.slide_layout.name!r}")
    dump_slide_shapes(slide, verbose)

    blocks_by_name, rescue_lines = parse_placeholder_blocks(body0_wo_layout)
    logger.debug(f"[page {page_no}] placeholder blocks: { {k: len(v) for k, v in blocks_by_name.items()} }")

    title_shape = find_title_shape(slide)

    # Build slide content based on layout type
    if layout_name == "Title Slide":
        build_title_slide(slide, front, body0_wo_layout, title_shape, verbose)
        allow_rescue = False
    else:
        rescue_lines = build_content_slide(
            slide, heading_text, heading_level, heading_raw, rescue_lines, title_shape
        )
        allow_rescue = True

    # Process placeholder blocks
    explicitly_used_shape_ids = process_placeholder_blocks(
        slide, blocks_by_name, md_dir, app_dir, state, page_no, verbose
    )

    # Process rescue content
    if allow_rescue:
        process_rescue_content(
            slide, rescue_lines, explicitly_used_shape_ids, state, page_no, verbose
        )


# -------------------------
# Main
# -------------------------
def main() -> int:
    """Main entry point."""
    parser = argparse.ArgumentParser(
        prog="conv-md-to-pptx.py",
        description="Convert Markdown to PowerPoint presentation"
    )
    parser.add_argument("src_md", help="Source markdown file")
    parser.add_argument("dst_pptx", help="Destination pptx")
    parser.add_argument(
        "--template",
        default="template.pptx",
        help='Template pptx path (default: "template.pptx")',
    )
    parser.add_argument("-v", "--verbose", action="store_true", help="Verbose logging to stdout")
    args = parser.parse_args()

    # Setup logging
    setup_logging(args.verbose)

    app_dir = get_app_dir()

    src_path = resolve_path(app_dir, args.src_md)
    dst_path = resolve_path(app_dir, args.dst_pptx)
    template_path = resolve_path(app_dir, args.template)

    # Validate input files
    if not src_path.exists():
        logger.error(
            f"ERROR: Source markdown file not found: {src_path}\n"
            f"  Current working directory: {Path.cwd()}\n"
            f"  App directory: {app_dir}\n"
            f"  Absolute path: {src_path.absolute()}"
        )
        return 2
    
    if not template_path.exists():
        logger.error(
            f"ERROR: Template pptx file not found: {template_path}\n"
            f"  Current working directory: {Path.cwd()}\n"
            f"  App directory: {app_dir}\n"
            f"  Absolute path: {template_path.absolute()}"
        )
        return 2

    # Read markdown
    try:
        md_text = src_path.read_text(encoding="utf-8")
    except (IOError, OSError, UnicodeDecodeError) as e:
        logger.error(f"ERROR: Failed to read markdown file: {e}")
        return 2

    pages = split_pages(md_text)
    logger.info(f"Processing {len(pages)} page(s)")

    # Load presentation template
    try:
        prs = Presentation(str(template_path))
    except Exception as e:
        logger.error(f"ERROR: Failed to load template: {e}")
        return 2

    state = {
        "text_written": {},  # track per (slide_id, placeholder_name) for append mode
        "verbose": args.verbose,
    }

    md_dir = src_path.parent

    # Build slides
    for page_no, page in enumerate(pages, start=1):
        if all(ln.strip() == "" for ln in page):
            continue
        try:
            build_slide_from_page(prs, page, md_dir=md_dir, app_dir=app_dir, state=state, page_no=page_no)
        except Exception as e:
            logger.error(f"ERROR: Failed to build slide {page_no}: {e}")
            if args.verbose:
                import traceback
                traceback.print_exc()
            return 2

    # Save output
    try:
        dst_path.parent.mkdir(parents=True, exist_ok=True)
        prs.save(str(dst_path))
        logger.info(f"Successfully created: {dst_path}")
    except (IOError, OSError) as e:
        logger.error(f"ERROR: Failed to save presentation: {e}")
        return 2

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
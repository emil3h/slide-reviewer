"""
slide_check.py - formatting checker with per-issue, individually-applyable fixes.

Rules (applied to every slide except the title slide, and skipping page-number/
heading/body checks on "Back-up Slides" dividers):
  1. Page number present, exactly one, Arial 8, numbered in order.
  2. Exactly one approved control marking (Not-CUI footer or CUI header banner),
     consistent deck-wide, worded exactly, formatted per its variant, no highlight.
  3. Logo present top-left (within 0.25"), width >= 3" (applies even on dividers).
  4. Heading (top-most text box) Arial 24 Bold, black.
  5. All other body text Arial, >= 16pt.

Public API used by app.py:
  review(prs) -> list[Issue]           # each Issue has .id, .slide, .message, .fixable
  apply_selected(prs, issues, ids)     # apply only the issues whose id is in ids

CLI:
  python slide_check.py deck.pptx            # report
  python slide_check.py deck.pptx --fix       # apply ALL fixable issues, write a copy
"""
import argparse
import re
import sys
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import PP_PLACEHOLDER, MSO_SHAPE_TYPE
from pptx.enum.dml import MSO_COLOR_TYPE
from lxml import etree

# ---- rules -----------------------------------------------------------------
HEADING_FONT, HEADING_SIZE, HEADING_BOLD = "Arial", 24.0, True
HEADING_COLOR = "000000"  # heading text must be black
BODY_FONT, BODY_SIZE = "Arial", 16.0
PAGENO_FONT, PAGENO_SIZE = "Arial", 8.0
PAGENO_ANCHOR_IN = (12.4, 6.95, 0.6, 0.3)
LOGO_MIN_WIDTH_IN = 3.0
LOGO_CORNER_TOLERANCE_IN = 0.25
CUI_VARIANTS = [
    {"id": "not_cui_footer", "label": "Not-CUI footer",
     "text": "Reviewed and determined not to contain CUI",
     "font": "Arial", "size": 8.0, "bold": True,
     "anchor_in": (4.42, 6.95, 4.5, 0.3)},
    {"id": "cui_header", "label": "CUI header banner",
     "text": "CUI//SP-EXPT",
     "font": "Arial", "size": 17.0, "bold": True,
     "anchor_in": (5.14, 0.0, 3.04, 0.4)},
]
SKIP_FIRST_SLIDE = True
MASTER_PROVIDES_PAGE_NUMBERS = False


@dataclass
class Issue:
    id: str
    slide: int
    category: str
    message: str
    fixable: bool
    fix: dict = field(default_factory=dict)


# ---- text / shape helpers --------------------------------------------------
def runs(shape):
    if not shape.has_text_frame:
        return []
    out = []
    for p in shape.text_frame.paragraphs:
        out.extend(p.runs)
    return out


def text_of(shape):
    return shape.text_frame.text.strip() if shape.has_text_frame else ""


def norm(s):
    return re.sub(r"\s+", " ", s.lower()).strip().rstrip(".")


def is_pageno(shape):
    if not shape.has_text_frame or shape.left is None or shape.top is None:
        return False
    if not re.fullmatch(r"\d{1,3}", text_of(shape)):
        return False
    return Emu(shape.left).inches > 7.0 and Emu(shape.top).inches > 6.0


def _ph_type(shape):
    """Get placeholder type or None."""
    try:
        return shape.placeholder_format.type if shape.is_placeholder else None
    except:
        return None

def is_slidenum(shape):
    return _ph_type(shape) == PP_PLACEHOLDER.SLIDE_NUMBER or \
           "slidenum" in _safe_xml(shape)

def is_number_text(shape):
    return shape.has_text_frame and bool(re.fullmatch(r"\d{1,3}", text_of(shape)))

def slidenum_has_value(shape):
    return is_slidenum(shape) and \
           ('type="slidenum"' in _safe_xml(shape) or any(c.isdigit() for c in text_of(shape)))

def is_footer_ph(shape):
    return _ph_type(shape) == PP_PLACEHOLDER.FOOTER

def is_title_ph(shape):
    return _ph_type(shape) == PP_PLACEHOLDER.TITLE

def is_body_ph(shape):
    pt = _ph_type(shape)
    return pt in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT, PP_PLACEHOLDER.SUBTITLE)


def is_pageno_carrier(shape):
    return is_pageno(shape) or is_number_text(shape) or is_slidenum(shape) \
        or is_footer_ph(shape)


def _safe_xml(shape):
    try:
        return shape._element.xml.lower()
    except Exception:
        return ""


def _slidenum_formatting(shape):
    """Extract explicit formatting overrides from slidenum placeholder XML."""
    try:
        root = etree.fromstring(shape._element.xml.encode('utf-8'))
        ns = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
        rPr = root.find(f".//{ns}rPr")
        if rPr is None:
            return {}

        result = {}
        if (latin := rPr.find(f"{ns}latin")) is not None:
            result['font'] = latin.get('typeface')
        if (sz := rPr.get('sz')):
            result['size'] = round(int(sz) / 100.0, 1)
        if (srgb := rPr.find(f".//{ns}solidFill/{ns}srgbClr")) is not None:
            result['color'] = srgb.get('val')
        return result
    except:
        return {}


def has_pageno(slide):
    if MASTER_PROVIDES_PAGE_NUMBERS:
        return True
    for s in slide.shapes:
        if is_pageno(s) or is_number_text(s):
            return True
        if is_slidenum(s) and slidenum_has_value(s):
            return True
        if is_footer_ph(s) and ("slidenum" in _safe_xml(s)
                                or any(c.isdigit() for c in text_of(s))):
            return True
    return False


def pageno_value(slide):
    for s in slide.shapes:
        if s.has_text_frame and re.fullmatch(r"\d{1,3}", text_of(s)):
            return int(text_of(s))
    return None


def literal_pageno_shape(slide):
    """A typed page-number text box (as opposed to an auto-updating slide-number
    field, whose cached text python-pptx can't safely rewrite)."""
    for s in slide.shapes:
        if is_number_text(s) and not is_slidenum(s):
            return s
    return None


def _has_slidenum_field(shape):
    return 'type="slidenum"' in _safe_xml(shape)


def _pageno_value_and_shape(slide):
    """Return (carrier_shape, numeric_value) for the page number *shown* on the
    slide, or (None, None). Reads the displayed digits from a typed number box,
    a literal slide-number placeholder, OR an auto slide-number field (whose
    rendered text may also include a stray typed digit, e.g. an auto "5" with an
    extra typed "5" showing as "55"). Digits are extracted from the full text so
    such mixed carriers are evaluated on what the viewer actually sees."""
    for s in slide.shapes:
        take = is_pageno(s) or is_number_text(s) or (is_slidenum(s) and slidenum_has_value(s))
        if not take:
            continue
        digits = "".join(c for c in text_of(s) if c.isdigit())
        if digits:
            return s, int(digits)
    return None, None


def _pageno_shapes(slide):
    return [s for s in slide.shapes
            if is_pageno(s)
            or (is_number_text(s) and not is_slidenum(s))
            or (is_slidenum(s) and slidenum_has_value(s))]


def _variant_by_id(vid):
    return next(v for v in CUI_VARIANTS if v["id"] == vid)


def matching_variant(shape):
    t = norm(text_of(shape))
    if not t:
        return None
    for v in CUI_VARIANTS:
        if t == norm(v["text"]):
            return v
    return None


def closest_variant(shape):
    t = text_of(shape)
    if not t:
        return None
    best, best_ratio = None, 0.0
    for v in CUI_VARIANTS:
        ratio = SequenceMatcher(None, norm(t), norm(v["text"])).ratio()
        if ratio > best_ratio:
            best, best_ratio = v, ratio
    return best if best_ratio > 0.7 else None


def is_marking_like(shape):
    return closest_variant(shape) is not None


def is_backup_divider(slide):
    texts = [text_of(s) for s in slide.shapes
             if s.has_text_frame and text_of(s)
             and not is_pageno_carrier(s) and not is_marking_like(s)]
    flat = re.sub(r"\s+", " ", " ".join(texts).lower().replace("-", " ")).strip()
    return flat in ("back up slides", "backup slides", "back up slide", "backup slide")


def is_logo(shape):
    if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
        return False
    if shape.left is None or shape.top is None or shape.width is None:
        return False
    return (Emu(shape.left).inches <= LOGO_CORNER_TOLERANCE_IN
            and Emu(shape.top).inches <= LOGO_CORNER_TOLERANCE_IN
            and Emu(shape.width).inches >= LOGO_MIN_WIDTH_IN)


def _rPr(run):
    return run._r.find(qn("a:rPr"))


def has_highlight(run):
    rPr = _rPr(run)
    return rPr is not None and rPr.find(qn("a:highlight")) is not None


def remove_highlight(run):
    rPr = _rPr(run)
    if rPr is not None:
        h = rPr.find(qn("a:highlight"))
        if h is not None:
            rPr.remove(h)


def inherited_shapes(slide):
    out = []
    try:
        layout = slide.slide_layout
        out.extend(layout.shapes)
        out.extend(layout.slide_master.shapes)
    except Exception:
        pass
    return out


def has_logo(slide):
    return any(is_logo(s) for s in slide.shapes) or \
        any(is_logo(s) for s in inherited_shapes(slide))


def _slide_marking_shapes(slide):
    out = {}
    for s in slide.shapes:
        v = matching_variant(s)
        if v and v["id"] not in out:
            out[v["id"]] = s
    return out


def _is_placeholder(shape):
    try:
        return shape.is_placeholder
    except Exception:
        return False


def _slide_marking_ids(slide):
    """Marking ids actually DISPLAYED on this slide.

    Counts (a) slide-level marking shapes — including an instantiated footer
    placeholder that carries the text — plus (b) inherited NON-placeholder
    shapes from the layout/master (e.g. the CUI header text box), which really
    do render on every slide.

    Inherited *placeholder prototypes* (a footer placeholder defined only on the
    master) are NOT counted: a placeholder renders on a slide only when that
    slide instantiates it, and that instance is already picked up by (a). Counting
    the master prototype made every slide look like it had the footer marking, so
    slides showing only the header were wrongly flagged "both present"."""
    ids = set(_slide_marking_shapes(slide))
    for s in inherited_shapes(slide):
        if _is_placeholder(s):
            continue
        v = matching_variant(s)
        if v:
            ids.add(v["id"])
    return ids


def _remove_variant_shapes(slide, variant):
    """Delete every shape matching `variant` from the slide, its layout, and its
    master. Removing from the master/layout affects every slide that shares it —
    that's the intended, deck-wide effect of resolving a "both present" conflict."""
    holders = [slide]
    try:
        layout = slide.slide_layout
        holders.append(layout)
        holders.append(layout.slide_master)
    except Exception:
        pass
    for holder in holders:
        for s in list(holder.shapes):
            v = matching_variant(s)
            if v is not None and v["id"] == variant["id"]:
                s._element.getparent().remove(s._element)


# ---- theme font-token resolution -------------------------------------------
_THEME_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
_A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
_theme_cache = {}
_master_pageno_cache = {}


def _theme_fonts(slide):
    try:
        master = slide.slide_layout.slide_master
    except Exception:
        return (None, None)
    key = id(master)
    if key in _theme_cache:
        return _theme_cache[key]
    major = minor = None
    try:
        blob = master.part.part_related_by(_THEME_RT).blob
        root = etree.fromstring(blob)
        mj = root.find(f".//{_A_NS}fontScheme/{_A_NS}majorFont/{_A_NS}latin")
        mn = root.find(f".//{_A_NS}fontScheme/{_A_NS}minorFont/{_A_NS}latin")
        major = mj.get("typeface") if mj is not None else None
        minor = mn.get("typeface") if mn is not None else None
    except Exception:
        pass
    _theme_cache[key] = (major, minor)
    return (major, minor)


def effective_font_name(name, slide):
    if name in ("+mj-lt", "+mj-ea", "+mj-cs"):
        return _theme_fonts(slide)[0]
    if name in ("+mn-lt", "+mn-ea", "+mn-cs"):
        return _theme_fonts(slide)[1]
    return name


def _master_slidenum(slide):
    """Get master's slidenum placeholder, cached."""
    try:
        master = slide.slide_layout.slide_master
        key = id(master)
        if key not in _master_pageno_cache:
            _master_pageno_cache[key] = next((s for s in master.shapes if is_slidenum(s)), None)
        return _master_pageno_cache[key]
    except:
        return None


def heading_shape(slide):
    for shp in slide.shapes:
        if is_title_ph(shp):
            return shp
    best, best_top = None, None
    for shp in slide.shapes:
        if not shp.has_text_frame or not text_of(shp):
            continue
        if is_pageno_carrier(shp) or is_marking_like(shp) or shp.top is None:
            continue
        top = Emu(shp.top).inches
        if best_top is None or top < best_top:
            best, best_top = shp, top
    return best


def explicit_rgb(run):
    """The run's explicitly-set RGB colour as a 6-hex string, or None when the
    colour is inherited or theme-linked (which we trust, to avoid false
    positives on text that simply follows the template default)."""
    try:
        c = run.font.color
        if c is not None and c.type == MSO_COLOR_TYPE.RGB:
            return str(c.rgb)
    except Exception:
        pass
    return None


# ---- review (produces Issues) ----------------------------------------------
def _font_issues(shape, want_font, want_size, where, slide, slide_no, target, idp, issues,
                  size_cmp="exact", want_bold=None, want_color=None):
    for r in runs(shape):
        name = effective_font_name(r.font.name, slide)
        if name is not None and name != want_font:
            issues.append(Issue(f"{idp}-font", slide_no, "font",
                                f"{where} font is {name}; should be {want_font}.", True,
                                {"op": "font_name", "value": want_font, **target}))
            break
    for r in runs(shape):
        if r.font.size is None:
            continue
        size = round(r.font.size.pt, 1)
        bad = size != want_size if size_cmp == "exact" else size < want_size
        if bad:
            rel = "exactly" if size_cmp == "exact" else "at least"
            issues.append(Issue(f"{idp}-size", slide_no, "size",
                                f"{where} size is {size}pt; should be {rel} "
                                f"{int(want_size)}pt.", True,
                                {"op": "font_size", "value": want_size, **target}))
            break
    if want_bold is not None:
        for r in runs(shape):
            # Flag if explicitly set to wrong value OR not set at all (None)
            if r.font.bold != want_bold:
                issues.append(Issue(f"{idp}-bold", slide_no, "bold",
                                    f"{where} is not bold; should be Bold.", True,
                                    {"op": "font_bold", "value": want_bold, **target}))
                break
    if want_color is not None:
        for r in runs(shape):
            rgb = explicit_rgb(r)
            if rgb is not None and rgb.upper() != want_color.upper():
                issues.append(Issue(f"{idp}-color", slide_no, "color",
                                    f"{where} font color is #{rgb}; should be black "
                                    f"(#{want_color}).", True,
                                    {"op": "font_color", "value": want_color, **target}))
                break


def _marking_format_defects(sh, variant, slide):
    defects = []
    rs = runs(sh)
    for r in rs:
        n = effective_font_name(r.font.name, slide)
        if n is not None and n != variant["font"]:
            defects.append(f"font {n}")
            break
    sizes = [round(r.font.size.pt, 1) for r in rs if r.font.size is not None]
    if sizes and any(sz != variant["size"] for sz in sizes):
        defects.append(f"not {int(variant['size'])}pt")
    bolds = [r.font.bold for r in rs if r.font.bold is not None]
    if bolds and not all(bolds):
        defects.append("not bold")
    if any(has_highlight(r) for r in rs):
        defects.append("has highlight")
    return defects


def _review_marking(slide, slide_no, canonical, issues):
    slide_variants = _slide_marking_shapes(slide)
    effective_ids = _slide_marking_ids(slide)
    marking_like = [s for s in slide.shapes
                    if is_marking_like(s) and matching_variant(s) is None]

    if len(effective_ids) > 1:
        issues.append(Issue(f"s{slide_no}-marking-both", slide_no, "marking",
                            "Both marking variants are present on this slide "
                            "(inherited from the shared slide master); exactly one "
                            "is required. Fixing this removes the other variant "
                            "from the master, which resolves it deck-wide.", True,
                            {"op": "resolve_marking_conflict", "keep": canonical["id"]}))
    elif not effective_ids:
        if marking_like:
            issues.append(Issue(f"s{slide_no}-marking-wording", slide_no, "marking",
                                f"Marking wording differs; must read exactly "
                                f"\u201c{canonical['text']}\u201d.", True,
                                {"op": "fix_marking_text", "shape_id": marking_like[0].shape_id,
                                 "variant": canonical["id"]}))
        else:
            issues.append(Issue(f"s{slide_no}-marking-missing", slide_no, "marking",
                                f"Missing required marking: \u201c{canonical['text']}\u201d.",
                                True, {"op": "add_marking", "variant": canonical["id"]}))
    else:
        vid = next(iter(effective_ids))
        if vid != canonical["id"]:
            sh = slide_variants.get(vid)
            issues.append(Issue(f"s{slide_no}-marking-variant", slide_no, "marking",
                                f"Marking uses the \u201c{_variant_by_id(vid)['label']}\u201d "
                                f"variant, but the deck's marking is "
                                f"\u201c{canonical['label']}\u201d.",
                                sh is not None,
                                {"op": "fix_marking_text", "shape_id": sh.shape_id,
                                 "variant": canonical["id"]} if sh is not None else {}))

    # Check for incorrectly-worded marking_like shapes (close but not exact)
    # even when a valid marking is present elsewhere on the slide or from master
    if marking_like and effective_ids:
        for sh in marking_like:
            issues.append(Issue(f"s{slide_no}-marking-wording-{sh.shape_id}",
                                slide_no, "marking",
                                f"Marking wording differs; must read exactly "
                                f"\u201c{canonical['text']}\u201d.", True,
                                {"op": "fix_marking_text", "shape_id": sh.shape_id,
                                 "variant": canonical["id"]}))

    # Format-check every slide-level marking shape present, independent of the
    # both-present branch above \u2014 a marking can be both duplicated *and*
    # incorrectly formatted (e.g. highlighted), and both need to be reported.
    for vid, sh in slide_variants.items():
        variant = _variant_by_id(vid)
        defects = _marking_format_defects(sh, variant, slide)
        if defects:
            issues.append(Issue(f"s{slide_no}-marking-format-{vid}", slide_no, "marking",
                                f"\u201c{variant['label']}\u201d marking formatting is off "
                                f"({', '.join(defects)}); should be {variant['font']} "
                                f"{int(variant['size'])}, bold, no highlight.", True,
                                {"op": "fix_marking_format", "shape_id": sh.shape_id,
                                 "variant": vid}))


def _review_slide(slide, slide_no, canonical, issues):
    divider = is_backup_divider(slide)

    # marking and logo apply to every non-title slide, including back-up dividers
    _review_marking(slide, slide_no, canonical, issues)

    if not has_logo(slide):
        issues.append(Issue(f"s{slide_no}-logo-missing", slide_no, "logo",
                            "No logo found in the top-left corner (min width 3\").",
                            False, {}))

    # back-up divider: skip page-number/heading/body checks
    if divider:
        return

    if not has_pageno(slide):
        issues.append(Issue(f"s{slide_no}-pageno-missing", slide_no, "page_number",
                            "No page number found on the slide.", True,
                            {"op": "add_pageno", "n": slide_no}))
    else:
        pn_shapes = _pageno_shapes(slide)
        if len(pn_shapes) > 1:
            issues.append(Issue(f"s{slide_no}-pageno-dup", slide_no, "page_number",
                                f"{len(pn_shapes)} page-number elements found; "
                                "exactly one is required.", False, {}))
        elif pn_shapes:
            sh = pn_shapes[0]
            # For slidenum placeholders, check position and formatting overrides
            if is_slidenum(sh):
                if (master := _master_slidenum(slide)) and sh.left and sh.top:
                    m_left, m_top = Emu(master.left).inches, Emu(master.top).inches
                    s_left, s_top = Emu(sh.left).inches, Emu(sh.top).inches
                    if abs(s_left - m_left) > 0.1 or abs(s_top - m_top) > 0.1:
                        issues.append(Issue(f"s{slide_no}-pageno-moved", slide_no, "page_number",
                                            f"Page number moved from master ({m_left:.2f}\", {m_top:.2f}\") "
                                            f"to ({s_left:.2f}\", {s_top:.2f}\").", False, {}))

                fmt = _slidenum_formatting(sh)
                if (font := fmt.get('font')) and font != PAGENO_FONT:
                    issues.append(Issue(f"s{slide_no}-pageno-font-override", slide_no, "page_number",
                                        f"Page number font overridden to {font}; should inherit {PAGENO_FONT}.",
                                        False, {}))
                if (size := fmt.get('size')) and size != PAGENO_SIZE:
                    issues.append(Issue(f"s{slide_no}-pageno-size-override", slide_no, "page_number",
                                        f"Page number size overridden to {size}pt; should inherit {int(PAGENO_SIZE)}pt.",
                                        False, {}))
            # Check position and font for typed number boxes
            elif is_number_text(sh) and not is_slidenum(sh):
                # Position check
                if sh.left is not None and sh.top is not None:
                    left_in = Emu(sh.left).inches
                    top_in = Emu(sh.top).inches
                    # Check if in bottom-right area (general check)
                    if left_in <= 7.0 or top_in <= 6.0:
                        issues.append(Issue(f"s{slide_no}-pageno-position", slide_no, "page_number",
                                            f"Page number position is ({left_in:.2f}\", {top_in:.2f}\"); "
                                            f"should be in bottom-right (left >7.0\", top >6.0\").",
                                            False, {}))
                # Font check
                _font_issues(sh, PAGENO_FONT, PAGENO_SIZE, "Page number", slide, slide_no,
                             {"target": "shape", "shape_id": sh.shape_id},
                             f"s{slide_no}-pageno", issues)

    head = heading_shape(slide)
    if head is None:
        issues.append(Issue(f"s{slide_no}-noheading", slide_no, "heading",
                            "No heading text box found.", False, {}))
    else:
        _font_issues(head, HEADING_FONT, HEADING_SIZE, "Heading", slide, slide_no,
                     {"target": "heading"}, f"s{slide_no}-heading", issues,
                     want_bold=HEADING_BOLD, want_color=HEADING_COLOR)

    for shp in slide.shapes:
        if not shp.has_text_frame:
            continue
        if is_pageno_carrier(shp) or is_marking_like(shp):
            continue
        if is_title_ph(shp):
            continue
        if head is not None and shp.shape_id == head.shape_id:
            continue
        # Check body placeholders even if empty, and all text boxes with content
        if is_body_ph(shp) or text_of(shp):
            _font_issues(shp, BODY_FONT, BODY_SIZE, "Body text", slide, slide_no,
                         {"target": "shape", "shape_id": shp.shape_id},
                         f"s{slide_no}-body-{shp.shape_id}", issues, size_cmp="min")


def detect_marking_variant(prs):
    """Best-guess deck-wide marking variant (majority vote across checked slides),
    for callers that want to show/default to it before the user overrides it."""
    slides = list(prs.slides)
    start = 2 if SKIP_FIRST_SLIDE else 1
    return _deck_marking_variant(slides, start)["id"]


def _deck_marking_variant(slides, start):
    counts = {v["id"]: 0 for v in CUI_VARIANTS}
    for idx, slide in enumerate(slides, start=1):
        if idx < start:
            continue
        for vid in _slide_marking_ids(slide):
            counts[vid] += 1
    best_id = max(counts, key=lambda vid: counts[vid])
    return _variant_by_id(best_id) if counts[best_id] > 0 else CUI_VARIANTS[0]


def _review_order(slides, start, issues):
    """Flag page numbers that don't match their sequential position.

    The expected number for a slide is its 1-based deck position plus an offset
    learned from the first slide that actually shows a number (so decks that
    start numbering at something other than 1 still work). This reads the value
    shown by literal boxes, literal slide-number placeholders, AND auto
    slide-number fields (including a field with a stray extra digit rendering as
    e.g. "55"), so out-of-order numbers in any carrier are caught."""
    offset = None
    for idx, slide in enumerate(slides, start=1):
        _, val = _pageno_value_and_shape(slide)
        if val is not None:
            offset = val - idx
            break
    if offset is None:
        return
    for idx, slide in enumerate(slides, start=1):
        if idx < start or is_backup_divider(slide):
            continue
        sh, val = _pageno_value_and_shape(slide)
        if val is None:
            continue
        expected = idx + offset
        if val != expected:
            issues.append(Issue(f"s{idx}-pageno-order", idx, "page_number",
                                f"Page number is {val} but should be {expected} "
                                f"for its position (out of order).", True,
                                {"op": "set_pageno_value", "shape_id": sh.shape_id,
                                 "value": expected}))


def review(prs, marking_variant=None):
    """marking_variant: id of the deck's intended marking (see CUI_VARIANTS) to
    mandate deck-wide, or None to auto-detect it from the majority of slides."""
    slides = list(prs.slides)
    start = 2 if SKIP_FIRST_SLIDE else 1
    issues = []
    canonical = _variant_by_id(marking_variant) if marking_variant \
        else _deck_marking_variant(slides, start)
    for idx, slide in enumerate(slides, start=1):
        if idx < start:
            continue
        _review_slide(slide, idx, canonical, issues)
    _review_order(slides, start, issues)
    issues.sort(key=lambda i: i.slide)
    return issues


# ---- fix primitives --------------------------------------------------------
def set_font(shape, name, size):
    for r in runs(shape):
        r.font.name = name
        r.font.size = Pt(size)


def set_text(shape, new_text):
    rs = runs(shape)
    if rs:
        rs[0].text = new_text
        for extra in rs[1:]:
            extra.text = ""
    elif shape.has_text_frame:
        shape.text_frame.paragraphs[0].add_run().text = new_text


def _set_pageno_value(shape, value):
    """Make the page number read `value`. If the carrier holds an auto
    slide-number field, drop any stray literal runs (the extra typed digit that
    made it show e.g. "55") and refresh the field's cached value, so the field
    keeps auto-numbering. Otherwise rewrite the literal text and re-apply the
    page-number font."""
    if not shape.has_text_frame:
        return
    if _has_slidenum_field(shape):
        for p in shape.text_frame.paragraphs:
            for r in list(p.runs):
                r._r.getparent().remove(r._r)
        for fld in shape._element.findall(".//" + qn("a:fld")):
            t = fld.find(qn("a:t"))
            if t is not None:
                t.text = str(value)
    else:
        set_text(shape, str(value))
        for r in runs(shape):
            r.font.name = PAGENO_FONT
            r.font.size = Pt(PAGENO_SIZE)


def _empty_slidenum_ph(slide):
    for s in slide.shapes:
        if is_slidenum(s) and not slidenum_has_value(s):
            return s
    return None


def add_pageno(slide, n):
    # Prefer filling an existing (empty) slide-number placeholder so the number
    # lands in the template's intended spot instead of a duplicate floating box.
    ph = _empty_slidenum_ph(slide)
    if ph is not None:
        set_text(ph, str(n))
        for r in runs(ph):
            r.font.name = PAGENO_FONT
            r.font.size = Pt(PAGENO_SIZE)
        return
    left, top, width, height = PAGENO_ANCHOR_IN
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    box.text_frame.margin_left = 0
    box.text_frame.margin_top = 0
    r = box.text_frame.paragraphs[0].add_run()
    r.text = str(n)
    r.font.name = PAGENO_FONT
    r.font.size = Pt(PAGENO_SIZE)
    r.font.color.rgb = RGBColor(0x88, 0x88, 0x88)


def add_marking_box(slide, variant):
    left, top, width, height = variant["anchor_in"]
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    box.text_frame.margin_top = 0
    box.text_frame.margin_bottom = 0
    p = box.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    p.add_run().text = variant["text"]
    return box


def format_marking_run(r, variant):
    r.font.name = variant["font"]
    r.font.size = Pt(variant["size"])
    r.font.bold = variant["bold"]
    remove_highlight(r)


def _shape_by_id(slide, sid):
    for s in slide.shapes:
        if s.shape_id == sid:
            return s
    return None


def _apply_fix(slide, fix):
    op = fix.get("op")
    if op == "font_name":
        sh = heading_shape(slide) if fix.get("target") == "heading" \
            else _shape_by_id(slide, fix.get("shape_id"))
        if sh:
            for r in runs(sh):
                r.font.name = fix["value"]
    elif op == "font_size":
        sh = heading_shape(slide) if fix.get("target") == "heading" \
            else _shape_by_id(slide, fix.get("shape_id"))
        if sh:
            for r in runs(sh):
                r.font.size = Pt(fix["value"])
    elif op == "font_bold":
        sh = heading_shape(slide) if fix.get("target") == "heading" \
            else _shape_by_id(slide, fix.get("shape_id"))
        if sh:
            for r in runs(sh):
                r.font.bold = fix["value"]
    elif op == "font_color":
        sh = heading_shape(slide) if fix.get("target") == "heading" \
            else _shape_by_id(slide, fix.get("shape_id"))
        if sh:
            for r in runs(sh):
                r.font.color.rgb = RGBColor.from_string(fix["value"])
    elif op == "add_pageno":
        add_pageno(slide, fix["n"])
    elif op == "set_number":
        sh = _shape_by_id(slide, fix.get("shape_id"))
        if sh:
            set_text(sh, str(fix["value"]))
    elif op == "set_pageno_value":
        sh = _shape_by_id(slide, fix.get("shape_id"))
        if sh:
            _set_pageno_value(sh, fix["value"])
    elif op == "add_marking":
        variant = _variant_by_id(fix["variant"])
        box = add_marking_box(slide, variant)
        for r in runs(box):
            format_marking_run(r, variant)
    elif op == "fix_marking_text":
        sh = _shape_by_id(slide, fix.get("shape_id"))
        variant = _variant_by_id(fix["variant"])
        if sh:
            set_text(sh, variant["text"])
            for r in runs(sh):
                format_marking_run(r, variant)
    elif op == "fix_marking_format":
        sh = _shape_by_id(slide, fix.get("shape_id"))
        variant = _variant_by_id(fix["variant"])
        if sh:
            for r in runs(sh):
                format_marking_run(r, variant)
    elif op == "resolve_marking_conflict":
        keep = fix.get("keep")
        for other in CUI_VARIANTS:
            if other["id"] != keep:
                _remove_variant_shapes(slide, other)


def apply_selected(prs, issues, ids):
    id_set = set(ids)
    slides = list(prs.slides)
    for i in issues:
        if i.id in id_set and i.fixable and 1 <= i.slide <= len(slides):
            _apply_fix(slides[i.slide - 1], i.fix)


# ---- CLI -------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("deck")
    ap.add_argument("--fix", action="store_true")
    ap.add_argument("--marking", choices=[v["id"] for v in CUI_VARIANTS],
                    help="Mandate the deck's approved marking variant instead of "
                         "auto-detecting it from the majority of slides.")
    args = ap.parse_args()
    path = Path(args.deck)
    if not path.exists():
        sys.exit(f"File not found: {path}")

    prs = Presentation(str(path))
    issues = review(prs, marking_variant=args.marking)
    skipped = " (slide 1 skipped as the title slide)" if SKIP_FIRST_SLIDE else ""
    if not issues:
        print(f"No issues found. All checked slides pass{skipped}.")
    else:
        print(f"\nFound {len(issues)} issue(s) in {path.name}{skipped}:")
        print("-" * 62)
        last = None
        for i in issues:
            if i.slide != last:
                print(f"\nSlide {i.slide}")
                last = i.slide
            tag = "" if i.fixable else " (manual)"
            print(f"  - {i.message}{tag}")
        print("-" * 62)

    if args.fix:
        apply_selected(prs, issues, [i.id for i in issues if i.fixable])
        out = path.with_name(f"corrected_{path.name}")
        prs.save(str(out))
        print(f"\nApplied all fixable issues. Wrote: {out}")
    elif issues:
        print("\nRun again with --fix to write a corrected copy.")


if __name__ == "__main__":
    main()

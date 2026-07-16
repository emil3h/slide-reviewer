"""
app.py - Streamlit interface for the Slide Reviewer (with per-issue checkboxes).

Run from the repo root:
    python3 -m streamlit run frontend/app.py
"""
import io
import sys
import tempfile
from itertools import groupby
from pathlib import Path

import streamlit as st
from PIL import Image, ImageChops
from pptx import Presentation

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import slide_check as sc

LOGO_PATH = Path(__file__).resolve().parent / "logo.png"


def _load_logo():
    """The source PNG is a small circular badge on a large, dark vignette
    background (mostly opaque, so a plain alpha-bbox crop barely trims it).
    Isolate the badge by luminance+alpha, crop tight, then flatten onto white
    so it renders crisp instead of tiny/blurry within the full canvas."""
    if not LOGO_PATH.exists():
        return None
    im = Image.open(LOGO_PATH).convert("RGBA")
    bright = im.convert("L").point(lambda p: 255 if p > 55 else 0)
    opaque = im.split()[3].point(lambda p: 255 if p > 10 else 0)
    bbox = ImageChops.multiply(bright, opaque).getbbox()
    if bbox:
        im = im.crop(bbox)
    flat = Image.new("RGB", im.size, "white")
    flat.paste(im, mask=im.split()[3])
    return flat


LOGO_IMAGE = _load_logo()

st.set_page_config(page_title="Slide Reviewer",
                   page_icon=LOGO_IMAGE if LOGO_IMAGE is not None else None,
                   layout="centered")

st.markdown(
    """
    <style>
    .stAlert {
        border-radius: 10px;
    }

    .slide-container {
        border: 1px solid #ddd;
        padding: 10px;
        border-radius: 10px;
        margin-bottom: 10px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.header("Review Rubric")
    st.markdown(
        """
        **Font Rules**
        - **Heading:** Arial, 24 pt, bold, black
        - **Body:** Arial, 16 pt or larger

        **Page Number**
        - Bottom-right, Arial, 8 pt
        - Sequential order

        **Logo**
        - Top-left
        - Width of at least 3 inches

        **Markings**
        - Exactly one variant
        - Exact wording and Arial formatting
        """
    )

def review_path(path, marking_variant=None):
    prs = Presentation(str(path))
    return prs, sc.review(prs, marking_variant=marking_variant)




header_logo, header_title = st.columns([1, 5], vertical_alignment="center")
if LOGO_IMAGE is not None:
    header_logo.image(LOGO_IMAGE, width=80)
header_title.title("Slide Reviewer")
st.caption("Checks each deck against the template rules (headings, body text, page "
           "numbers, markings) and lets you approve fixes one by one.")

# ============================ SINGLE DECK ===================================
uploaded = st.file_uploader("Upload a presentation (.pptx)", type="pptx")
if uploaded is None:
    st.info("Upload a .pptx to begin.")
    st.stop()

# persist the upload to a temp file so it survives Streamlit reruns
if st.session_state.get("deck_name") != uploaded.name:
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
    tmp.write(uploaded.getbuffer())
    tmp.flush()
    st.session_state["deck_path"] = tmp.name
    st.session_state["deck_name"] = uploaded.name
    st.session_state["marking_choice"] = "auto"  # reset override for the new deck

variant_labels = {v["id"]: v["label"] for v in sc.CUI_VARIANTS}
detected = sc.detect_marking_variant(Presentation(st.session_state["deck_path"]))
choice = st.selectbox(
    "This deck's approved marking",
    ["auto"] + list(variant_labels),
    format_func=lambda vid: f"Auto-detect (currently: {variant_labels[detected]})"
                              if vid == "auto" else variant_labels[vid],
    key="marking_choice",
    help="The deck must use exactly one marking variant throughout. Pick it once "
         "here and every slide is checked against it.")
marking_variant = None if choice == "auto" else choice

prs, issues = review_path(st.session_state["deck_path"], marking_variant)
n_slides = len(list(prs.slides))
st.subheader(f"{uploaded.name} \u2014 {n_slides} slides")

if not issues:
    st.success("No issues found. This deck passes all checks.")
    st.stop()

fixable_ids = [i.id for i in issues if i.fixable]
for iid in fixable_ids:
    st.session_state.setdefault(iid, True)

n_fixable = len(fixable_ids)
n_manual = len(issues) - n_fixable

m1, m2, m3 = st.columns(3)
m1.metric("Issues found", len(issues))
m2.metric("Auto-fixable", n_fixable)
m3.metric("Manual", n_manual)
st.caption("Untick anything below you don't want changed.")

c1, c2, c3 = st.columns([1, 1, 2])
if c1.button("Select all", use_container_width=True):
    for iid in fixable_ids:
        st.session_state[iid] = True
    st.toast("All fixable issues selected")
if c2.button("Clear all", use_container_width=True):
    for iid in fixable_ids:
        st.session_state[iid] = False
    st.toast("All selections cleared")

hide_compliant = c3.checkbox("Hide compliant slides", value=False)

st.divider()

# render checkboxes grouped by slide, one card per slide
for slide_no, slide_issues in groupby(issues, key=lambda i: i.slide):
    slide_issues = list(slide_issues)
    
    if hide_compliant and not slide_issues:
        continue
        
    with st.container(border=True):
        st.markdown(f"**Slide {slide_no}**")
        for i in slide_issues:
            if i.fixable:
                st.checkbox(i.message, key=i.id)
            else:
                st.warning(f"**Manual Fix Required:** {i.message}")

st.divider()
selected = [i.id for i in issues if i.fixable and st.session_state.get(i.id)]
if st.button(f"Apply {len(selected)} selected fix(es)", type="primary",
             disabled=not selected, use_container_width=True):
    with st.spinner("Applying fixes..."):
        sc.apply_selected(prs, issues, selected)
        buf = io.BytesIO()
        prs.save(buf)
    st.success(f"Applied {len(selected)} fix(es).")
    st.download_button("Download corrected deck", data=buf.getvalue(),
                       file_name=f"corrected_{uploaded.name}",
                       mime="application/vnd.openxmlformats-officedocument."
                             "presentationml.presentation")


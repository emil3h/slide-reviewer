"""
app.py - Streamlit interface for the Slide Reviewer (with per-issue checkboxes).

Run from the repo root:
    python3 -m streamlit run frontend/app.py
"""
import io
import sys
import zipfile
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


def review_path(path, marking_variant=None):
    prs = Presentation(str(path))
    return prs, sc.review(prs, marking_variant=marking_variant)


def fix_all_bytes(path):
    prs = Presentation(str(path))
    issues = sc.review(prs)
    sc.apply_selected(prs, issues, [i.id for i in issues if i.fixable])
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


header_logo, header_title = st.columns([1, 5], vertical_alignment="center")
if LOGO_IMAGE is not None:
    header_logo.image(LOGO_IMAGE, width=80)
header_title.title("Slide Reviewer")
st.caption("Checks each deck against the template rules (headings, body text, page "
           "numbers, markings) and lets you approve fixes one by one.")

mode = st.radio("Mode", ["Single deck (upload)", "Batch (local folder)"], horizontal=True)

# ============================ SINGLE DECK ===================================
if mode == "Single deck (upload)":
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

    c1, c2 = st.columns(2)
    if c1.button("Select all", use_container_width=True):
        for iid in fixable_ids:
            st.session_state[iid] = True
    if c2.button("Clear all", use_container_width=True):
        for iid in fixable_ids:
            st.session_state[iid] = False

    st.divider()

    # render checkboxes grouped by slide, one card per slide
    for slide_no, slide_issues in groupby(issues, key=lambda i: i.slide):
        slide_issues = list(slide_issues)
        with st.container(border=True):
            st.markdown(f"**Slide {slide_no}**")
            for i in slide_issues:
                if i.fixable:
                    st.checkbox(i.message, key=i.id)
                else:
                    st.checkbox(i.message, value=False, disabled=True, key=i.id)
                    st.caption("Needs a manual fix in PowerPoint.")

    st.divider()
    selected = [i.id for i in issues if i.fixable and st.session_state.get(i.id)]
    if st.button(f"Apply {len(selected)} selected fix(es)", type="primary",
                 disabled=not selected, use_container_width=True):
        sc.apply_selected(prs, issues, selected)
        buf = io.BytesIO()
        prs.save(buf)
        st.success(f"Applied {len(selected)} fix(es).")
        st.download_button("Download corrected deck", data=buf.getvalue(),
                           file_name=f"corrected_{uploaded.name}",
                           mime="application/vnd.openxmlformats-officedocument."
                                "presentationml.presentation")

# ============================ BATCH FOLDER ==================================
else:
    folder = st.text_input("Folder containing .pptx decks", value="submitted_decks")
    if not folder:
        st.stop()
    path = Path(folder)
    if not path.exists():
        st.error(f"Folder not found: {path.resolve()}")
        st.stop()
    decks = sorted(path.glob("*.pptx"))
    if not decks:
        st.warning("No .pptx files found there.")
        st.stop()

    st.caption(f"Found {len(decks)} deck(s) in this folder.")
    if st.button("Review all decks", type="primary", use_container_width=True):
        rows = []
        progress = st.progress(0.0)
        for i, deck in enumerate(decks, start=1):
            _, issues = review_path(deck)
            rows.append({"Deck": deck.name, "Issues": len(issues)})
            progress.progress(i / len(decks))
        st.session_state["batch_rows"] = rows
        st.session_state["batch_folder"] = str(path)

    if "batch_rows" in st.session_state:
        rows = st.session_state["batch_rows"]
        total = sum(r["Issues"] for r in rows)
        st.subheader("Summary")
        m1, m2 = st.columns(2)
        m1.metric("Decks reviewed", len(rows))
        m2.metric("Total issues", total)
        st.dataframe(rows, use_container_width=True, hide_index=True)
        st.divider()
        if st.button("Apply all fixes to every deck & download zip",
                     use_container_width=True):
            zbuf = io.BytesIO()
            fp = Path(st.session_state["batch_folder"])
            with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as zf:
                for deck in sorted(fp.glob("*.pptx")):
                    zf.writestr(f"corrected_{deck.name}", fix_all_bytes(deck))
            st.success("All decks corrected.")
            st.download_button("Download all corrected decks (zip)",
                               data=zbuf.getvalue(), file_name="corrected_decks.zip",
                               mime="application/zip")

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
from pptx import Presentation

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import slide_check as sc

LOGO_PATH = Path(__file__).resolve().parent / "logo.png"
LOGO_IMAGE = str(LOGO_PATH) if LOGO_PATH.exists() else None

st.set_page_config(page_title="Slide Reviewer",
                   page_icon=LOGO_IMAGE if LOGO_IMAGE is not None else None,
                   layout="centered")


def inject_custom_styles():
    """Purely cosmetic: card elevation, tighter alert/metric chrome, and
    hover/transition polish on the buttons tagged with `key=` below. `.streamlit/
    config.toml` owns the actual color/font theme; this only fills in what
    config.toml can't reach."""
    st.html("""
    <style>
        .block-container {
            padding-top: 2rem !important;
            padding-bottom: 3rem !important;
            max-width: 900px !important;
        }

        /* st.container(border=True) -> elevated card */
        div[data-testid="stVerticalBlockBorderWrapper"] {
            border-radius: 14px !important;
            box-shadow: 0 1px 3px rgba(15, 23, 42, 0.06), 0 1px 2px rgba(15, 23, 42, 0.04) !important;
            transition: box-shadow 0.2s ease !important;
        }
        div[data-testid="stVerticalBlockBorderWrapper"]:hover {
            box-shadow: 0 6px 16px rgba(15, 23, 42, 0.08) !important;
        }

        div[data-testid="stAlertContainer"] {
            border-radius: 10px !important;
        }

        div[data-testid="stMetric"] {
            background: rgba(37, 99, 235, 0.05);
            border-radius: 10px;
            padding: 0.75rem 1rem;
        }

        /* Primary CTA */
        .st-key-apply-fixes-btn button {
            background-color: #2563EB !important;
            color: white !important;
            border-radius: 8px !important;
            border: none !important;
            font-weight: 600 !important;
            transition: all 0.15s ease !important;
        }
        .st-key-apply-fixes-btn button:hover {
            background-color: #1D4ED8 !important;
            transform: translateY(-1px);
            box-shadow: 0 4px 10px rgba(37, 99, 235, 0.35) !important;
        }

        /* Toolbar buttons */
        .st-key-select-all-btn button,
        .st-key-clear-all-btn button {
            border-radius: 8px !important;
            transition: all 0.15s ease !important;
        }
    </style>
    """)


inject_custom_styles()

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




header_logo, header_title = st.columns([1, 4], vertical_alignment="center")
if LOGO_IMAGE is not None:
    header_logo.image(LOGO_IMAGE, width=140)
header_title.title("Slide Reviewer")
st.caption("Checks each deck against the template rules (headings, body text, page "
           "numbers, markings) and lets you approve fixes one by one.")

# ============================ SINGLE DECK ===================================
with st.container(border=True):
    st.subheader("Upload deck")
    uploaded = st.file_uploader("Upload a presentation (.pptx)", type="pptx")
    if uploaded is None:
        st.info("Upload a .pptx to begin.", icon=":material/upload_file:")
        st.stop()

# persist the upload to a temp file so it survives Streamlit reruns
if st.session_state.get("deck_name") != uploaded.name:
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
    tmp.write(uploaded.getbuffer())
    tmp.flush()
    st.session_state["deck_path"] = tmp.name
    st.session_state["deck_name"] = uploaded.name
    # Auto-detect marking for the new deck
    detected = sc.detect_marking_variant(Presentation(st.session_state["deck_path"]))
    st.session_state["marking_choice"] = detected
elif "marking_choice" not in st.session_state:
    detected = sc.detect_marking_variant(Presentation(st.session_state["deck_path"]))
    st.session_state["marking_choice"] = detected

marking_variant = st.session_state["marking_choice"]
with st.spinner("Analyzing deck...", show_time=True):
    prs, issues = review_path(st.session_state["deck_path"], marking_variant)
n_slides = len(list(prs.slides))
st.subheader(f"{uploaded.name} — {n_slides} slides")

if not issues:
    st.success("No issues found. This deck passes all checks.", icon=":material/check_circle:")
    st.stop()

fixable_ids = [i.id for i in issues if i.fixable]
for iid in fixable_ids:
    st.session_state.setdefault(iid, True)

n_fixable = len(fixable_ids)
n_manual = len(issues) - n_fixable

with st.container(border=True):
    st.subheader("Overview")
    m1, m2, m3 = st.columns(3)
    m1.metric("Issues found", len(issues))
    m2.metric("Auto-fixable", n_fixable)
    m3.metric("Manual", n_manual)
    st.caption("Untick anything below you don't want changed.")

    c1, c2, c3 = st.columns([1, 1, 2])
    if c1.button("Select all", key="select-all-btn", use_container_width=True):
        for iid in fixable_ids:
            st.session_state[iid] = True
        st.toast("All fixable issues selected")
    if c2.button("Clear all", key="clear-all-btn", use_container_width=True):
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
                st.warning(f"**Manual Fix Required:** {i.message}", icon=":material/build:")

st.divider()
selected = [i.id for i in issues if i.fixable and st.session_state.get(i.id)]
custom_labels = {v["id"]: v["label"] for v in sc.CUI_VARIANTS}
# Override with user-requested labels
custom_labels["not_cui_footer"] = "Reviewed and determined not to contain CUI"
custom_labels["cui_header"] = "CUI header banner"

st.selectbox(
    "Approved marking for this deck",
    options=list(custom_labels.keys()),
    format_func=lambda x: custom_labels[x],
    key="marking_choice",
    help="The deck must use exactly one marking variant throughout. Changing this will re-run the review.",
)

if st.button(f"Apply {len(selected)} selected fix(es)", key="apply-fixes-btn", type="primary",
             disabled=not selected, use_container_width=True):
    with st.spinner("Applying fixes..."):
        sc.apply_selected(prs, issues, selected)
        buf = io.BytesIO()
        prs.save(buf)
    st.toast(f"Applied {len(selected)} fix(es)", icon=":material/check_circle:")
    st.success(f"Applied {len(selected)} fix(es).", icon=":material/check_circle:")
    st.download_button("Download corrected deck", data=buf.getvalue(),
                       file_name=f"corrected_{uploaded.name}",
                       mime="application/vnd.openxmlformats-officedocument."
                             "presentationml.presentation")

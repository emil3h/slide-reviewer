# 🛝 Slide Reviewer

Slide Reviewer is a tool designed to automatically check PowerPoint presentations (`.pptx`) against specific formatting and compliance rules. It can identify issues with page numbering, control markings, logo placement, and font consistency, and can automatically fix many of these issues..

## 🚀 Features

- **Automated Compliance Checks**:
    - **Page Numbers**: Verifies that exactly one page number is present, in the correct numerical order, and formatted Arial 8pt.
    - **Markings**: Ensures exactly one of the two approved control markings is present, worded exactly, formatted correctly, and consistent across the deck.
    - **Logo**: Verifies a logo is present in the top-left corner and at least 3" wide.
    - **Font Consistency**:
        - Headings: Checks for Arial, exactly 24pt, Bold.
        - Body Text: Checks for Arial, at least 16pt.
- **Interactive Web Interface**: A Streamlit-based UI to upload decks, review issues per slide, and selectively apply fixes.
- **Batch Processing**: Ability to review an entire folder of presentations and download all corrected versions in a single ZIP file.
- **CLI Tool**: A lightweight command-line interface for quick reports and bulk fixing.

## 🛠️ Installation

### Prerequisites
- Python 3.x
- `python-pptx`
- `streamlit` (for the web interface)
- `lxml` (for theme font resolution)

Install dependencies via pip:
```bash
pip install python-pptx streamlit lxml
```

## 📖 Usage

### Web Interface (Streamlit)
The easiest way to use the tool is via the Streamlit app (run from the repo root):
```bash
python3 -m streamlit run frontend/app.py
```
(On Windows, use `python` or `py` in place of `python3`.)
1. **Single Deck**: Upload a `.pptx` file. The app will list all issues. Check the boxes for the fixes you want to apply, then click "Apply selected fix(es)" and download the result.
2. **Batch Mode**: Enter the path to a folder containing multiple `.pptx` files. The app will summarize the issues across all decks and allow you to apply all auto-fixable issues to all decks at once.

### Command Line Interface (CLI)
You can use `slide_check.py` directly from the terminal:

**To report issues:**
```bash
python3 slide_check.py your_presentation.pptx
```

**To report and automatically apply all fixable issues:**
```bash
python3 slide_check.py your_presentation.pptx --fix
```
This will create a new file named `corrected_your_presentation.pptx`.

## ⚙️ Rules & Logic

The tool applies the following rules to every slide except the title slide. On
"Back-up Slides" divider slides, only the **Logo** and **Marking** checks apply; page
number, heading, and body checks are skipped.

| Element | Required Format | Auto-Fixable? |
| :--- | :--- | :--- |
| **Page Number** | Present, exactly one, bottom-right, Arial 8pt, sequential | Yes (missing/out-of-order); duplicates are flagged for manual removal |
| **Marking** | Exactly one of two approved variants, deck-wide consistent: "Reviewed and determined not to contain CUI" (footer, Arial 8, Bold, no highlight) or "CUI//SP-EXPT" (header, Arial 17, Bold, no highlight) | Yes — including "both present," which removes the unused variant from the slide master (you choose which to keep) |
| **Logo** | Picture, top-left corner (within 0.25"), width ≥ 3" | No — flagged for manual fix |
| **Heading** | Arial, exactly 24pt, Bold | Yes |
| **Body Text** | Arial, at least 16pt (larger is fine) | Yes |

Elements correctly provided by the slide master (logo, marking, page number) count as
present/compliant for slides that don't override them — only an explicit, non-compliant
override on the slide itself is flagged.

## 📂 Project Structure
- `frontend/app.py`: Streamlit web application.
- `slide_check.py`: Core analysis and fixing engine.
- `formatting_rules.json`: Configuration for formatting rules.
- `human_review_rubric.md`: Prose rubric the rules are derived from.
- `data/`: Reference/test decks (`template.pptx`, `testTemplate1.pptx`).

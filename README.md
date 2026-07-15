# 🛝 Slide Reviewer

Slide Reviewer is a tool designed to automatically check PowerPoint presentations (`.pptx`) against specific formatting and compliance rules. It can identify issues with page numbering, CUI markings, and font consistency, and can automatically fix many of these issues.

## 🚀 Features

- **Automated Compliance Checks**:
    - **Page Numbers**: Verifies that page numbers are present and in the correct numerical order.
    - **CUI Markings**: Ensures the required "Reviewed and determined not to contain CUI" marking is present, worded exactly, and formatted correctly (Arial 8, Bold, no highlight).
    - **Font Consistency**:
        - Headings: Checks for Arial 24pt.
        - Body Text: Checks for Arial 16pt.
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
The easiest way to use the tool is via the Streamlit app:
```bash
py -m streamlit run app.py
```
1. **Single Deck**: Upload a `.pptx` file. The app will list all issues. Check the boxes for the fixes you want to apply, then click "Apply selected fix(es)" and download the result.
2. **Batch Mode**: Enter the path to a folder containing multiple `.pptx` files. The app will summarize the issues across all decks and allow you to apply all auto-fixable issues to all decks at once.

### Command Line Interface (CLI)
You can use `slide_check.py` directly from the terminal:

**To report issues:**
```bash
python slide_check.py your_presentation.pptx
```

**To report and automatically apply all fixable issues:**
```bash
python slide_check.py your_presentation.pptx --fix
```
This will create a new file named `corrected_your_presentation.pptx`.

## ⚙️ Rules & Logic

The tool applies the following rules to every slide (except the title slide and specific "Back-up Slides" dividers):

| Element | Required Format | Auto-Fixable? |
| :--- | :--- | :--- |
| **Page Number** | Present and sequential | Yes |
| **CUI Marking** | "Reviewed and determined not to contain CUI" (Arial 8, Bold, no highlight) | Yes |
| **Heading** | Arial 24pt | Yes |
| **Body Text** | Arial 16pt | Yes |

## 📂 Project Structure
- `app.py`: Streamlit web application.
- `slide_check.py`: Core analysis and fixing engine.
- `formatting_rules.json`: Configuration for formatting rules.
- `template.pptx`: Reference template for correct formatting.

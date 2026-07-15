# Slide Reviewer

Checks PowerPoint decks (`.pptx`) against formatting/compliance rules (headings, body
text, page numbers, CUI markings) and can auto-fix issues. Streamlit UI + CLI.

## Run

```bash
pip install python-pptx streamlit lxml
python3 -m streamlit run frontend/app.py # web UI (run from repo root)
python3 slide_check.py deck.pptx         # CLI report
python3 slide_check.py deck.pptx --fix   # CLI report + write corrected_deck.pptx
```
(On Windows, use `python` or `py` in place of `python3`.)

## Structure

- `frontend/app.py` — Streamlit UI (single-deck upload + batch-folder modes). Thin
  wrapper around `slide_check`; inserts the repo root onto `sys.path` to import it.
  `frontend/logo.png` is shown in the app header.
- `slide_check.py` — core engine (repo root). Public API: `review(prs) -> list[Issue]`
  and `apply_selected(prs, issues, ids)`. This is where rule logic and fixes live.
- `formatting_rules.json` — declarative rule config (fonts, sizes, positions,
  CUI marking variants, backup-slide handling). `human_review_rubric.md` is the
  prose source these rules were derived from — keep them in sync if either changes.
- `data/` — reference/test decks (`template.pptx`, `testTemplate1.pptx`).

## Notes

- Rules skip the title slide and "Back-up Slides" divider slides (only logo/marking
  checks apply there per `formatting_rules.json`'s `sections.backup_slides`).
- Each `Issue` has a stable `.id` so the UI can let users approve/reject fixes
  individually rather than all-or-nothing.

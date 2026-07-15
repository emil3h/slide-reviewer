# Human Review Rubric — Slide Formatting Compliance

**Scope:** This rubric governs **formatting** compliance only (not content quality). A
reviewer — or the Slide Reviewer tool acting on their behalf — walks each slide and marks
every item below as **pass** or **deviation**. Deviations are presented for one-by-one
approval before any fix is applied. All data in this project is mock data.

**Deck baseline:** Widescreen, 13.33" × 7.5". The reference template carries the logo, both
marking variants, and the slide-number field **in the slide master**, so those elements are
normally *inherited* rather than placed on each slide (see [Inheritance & the master](#inheritance--the-master)).

**Title slide is fully exempt.** The first slide (the title slide) is **not checked against any
rule** — fonts, page number, logo, and marking are all skipped for it. Every rule below applies
to the remaining slides.

---

## 1. Fonts

Applies to every slide except the [title slide](#title-slide-fully-exempt) and
[back-up slides](#5-back-up-slides).

| Element | Typeface | Size | Weight |
|---|---|---|---|
| **Heading** (the slide's title / top-most heading text box) | Arial | **exactly 24 pt** | **Bold** |
| **Body** (all other body text) | Arial | **≥ 16 pt** (minimum, larger is fine) | — |

Rules:

- **Every** run of text on the slide must be **Arial** — headings, body, and markings alike.
  (Marking and page-number sizes are governed by their own sections below.)
- The **heading** must be Arial **24 pt Bold** — 24 is exact, and bold is required.
- **Body** text must be Arial and **at least 16 pt**. Sizes above 16 pt pass; anything under
  16 pt is a deviation.
- Font names that resolve through the theme (e.g. `+mn-lt`, `+mj-lt`) must resolve to Arial;
  because this template's theme is Aptos, theme-linked text will read as a deviation unless an
  explicit Arial override is set. See [Theme-linked fonts](#theme-linked-fonts).

## 2. Page number

- Every slide has a **page number** in the **bottom-right** corner.
- **Exactly one** page-number element per slide — no duplicates.
- Page number is **Arial 8 pt**.
- Page numbers must be **sequential / in order** across the deck (a number out of sequence
  for its position is a deviation).

## 3. Logo

- A **logo** appears in the **top-left corner**, **aligned to the corner** (flush to the
  top and left edges, within a small tolerance).
- The logo's **width is at least 3"**.
- The logo may be inherited from the master (the reference template places it there); an
  inherited logo counts as present.

## 4. Marking language

Each slide must carry a control marking. **Exactly one** of the two approved markings below
applies to a given slide, and it must match the **exact wording** and the **exact formatting**:

| Variant | Exact text | Placement | Formatting |
|---|---|---|---|
| **Not-CUI footer** | `Reviewed and determined not to contain CUI` | Footer (bottom) | Arial **8 pt**, **Bold**, **no highlight** |
| **CUI header banner** | `CUI//SP-EXPT` | Header (top) | Arial **17 pt**, **Bold**, **no highlight** |

Rules:

- **Exactly one** approved marking must be **displayed on every slide** — at least one and
  only one:
  - **Zero** markings on a slide → deviation (**missing**).
  - **Both** the footer *and* the header showing → deviation (**both present**).
- **Consistency across the deck:** all slides must use the **same** marking variant. The choice
  is made **once, deck-wide, in the master** — the author keeps the marking the deck needs and
  **deletes the other from the master**, so every slide inherits the same one. A slide showing a
  *different* variant than the rest of the deck is a deviation. The template intentionally ships
  **both** markings so the author can decide; exactly one survives after the master is edited.
- Wording must match **exactly** (a marking with different wording is a deviation, not a pass).
- Formatting is checked **strictly**: wrong typeface, wrong size, missing bold, or any text
  highlight is a deviation.
- Location: the marking may be a **text box on the slide** or **inherited from the master** —
  **the master is a valid location**, and a marking inherited from the master counts toward the
  "displayed on every slide" tally. (Reviewer note: the intent is explicitly to check the
  master, not only slide-level text boxes.)

## 5. Back-up slides

- A slide is a **back-up slide** if its heading/text reads **"Back-up Slides"**.
- On back-up slides, **only** check **marking language** (§4) and **logo placement** (§3).
- **Skip** the font checks (§1) and the page-number check (§2) on back-up slides.

---

## Judgment rules

### Inheritance & the master

Elements defined on the **slide master/layout** are inherited by every slide. For this
rubric, an element that is correctly provided by the master (logo, marking, slide number)
counts as **present and compliant** for slides that don't override it. Only an **explicit,
non-compliant override on the slide itself** is a deviation. This avoids false positives on
text that simply inherits correct defaults.

### Theme-linked fonts

PowerPoint stores theme fonts as tokens (`+mj-lt` = heading, `+mn-lt` = body). These must be
**resolved to the real typeface via the theme** before comparing to "Arial." This template's
theme resolves to **Aptos / Aptos Display**, so any text relying on the theme (rather than an
explicit Arial setting) will resolve to Aptos and read as a **font deviation**.

### Title slide (fully exempt)

The **first slide** is treated as the title slide and is **exempt from every rule** — no font,
page-number, logo, or marking checks are applied to it. All rules begin at the second slide.

---

## Reviewer checklist (per slide)

For each slide **except the title slide** (fully exempt), confirm:

- [ ] Heading is Arial, 24 pt, Bold *(skip on back-up slides)*
- [ ] All body text is Arial, ≥ 16 pt *(skip on back-up slides)*
- [ ] Exactly one page number, bottom-right, Arial 8 pt, in sequential order *(skip on back-up slides)*
- [ ] Logo present top-left, corner-aligned, width ≥ 3"
- [ ] Exactly one approved marking displayed — neither zero nor both — matching the deck's single chosen variant, exact wording, exact formatting

---

## Decisions log

All interpretation questions have been resolved; the rules above reflect them.

- **Marking count = exactly one.** Each slide must show **at least one and only one** approved
  marking. Zero → missing deviation; both → deviation. *(§4)*
- **Marking is one variant, deck-wide, on every slide.** The tally **counts markings displayed
  on the slide, including those inherited from the master.** Every slide must show exactly one,
  and all slides must use the **same** variant — the choice is made once by deleting the unused
  marking from the master. The checker reads the deck's intended variant from whichever marking
  remains in the master, and flags any slide that is missing a marking, shows both, or uses the
  off-variant. *(§4)*
- **Page-number order = kept.** Sequential/in-order numbering remains a checked requirement. *(§2)*
- **Title slide = fully exempt.** The first slide is not checked against any rule. *(Scope /
  [Title slide](#title-slide-fully-exempt))*
- **Back-up scope = the single divider slide only.** The exemption applies to the slide whose
  text reads "Back-up Slides," not to subsequent slides. *(§5)*
- **Logo identification.** A **picture** shape whose top-left is within **≤ 0.25"** of the
  (0, 0) corner and whose **width ≥ 3"**; a master-level logo satisfies the check for all
  slides. *(§3)*

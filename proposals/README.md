# RoboAgentix — AWNAK proposal documents

Two Arabic (RTL) proposal documents for **عَوْنَك / AWNAK**, the home-and-family
services platform, produced for RoboAgentix For Software Development.

Each is produced as both a **PDF** (the presentation artefact) and a **.docx**
(the working one — prices and wording stay editable).

| Output | Pages | Covers |
| --- | --- | --- |
| `RoboAgentix-AWNAK-01-Technical-Proposal-AR` | 34 | What gets built and with which technology — architecture, stack, data model, dispatch engine, order lifecycle, money layer, integrations, security, performance, delivery artefacts |
| `RoboAgentix-AWNAK-02-Commercial-Proposal-AR` | 22 | Line-item pricing, delivery model, 16-week schedule, payment plan, warranty and support, change control, ownership, acceptance criteria |

Headline commercial figures: **$5,800** for phase one over **16 weeks**, five
payments tied to milestones, 180 days of warranty, and **$4,700** of phase-two
options priced individually. Each external integration is priced on its own
line, as the specification requires.

Document references are `RAX-TEC-AWNAK-001` and `RAX-FIN-AWNAK-001`, both Rev. A.
The two are written to be read together: every price line in the commercial
document points at a scope described in the technical one.

## Sources

- **Requirements** — the client's 32-section Arabic specification
  (*وثيقة توصيف تطبيق خدمات المنزل والعائلة*). Every section of it is answered
  somewhere in the technical document.
- **Product design** — the approved AWNAK customer-app prototype. Its 15 screens
  were captured into `assets/screens/` and its design tokens (teal `#147D73`,
  Tajawal) are documented in §12 of the technical proposal. The prototype is
  what the client is shown; no Figma files are part of the deliverables.
- **Document design** — the RoboAgentix house proposal style: dark forest-green
  cover, black section headings over a green rule, green table headers, stat
  tiles, and a running header/footer carrying the logo, the revision line, the
  company contact line, and a page number.

## Build

```bash
pip install playwright pypdfium2 pillow python-docx beautifulsoup4 lxml
python3 proposals/build.py          # PDFs   — both documents
python3 proposals/build_docx.py     # Word   — both documents
python3 proposals/build.py 01       # just the technical one
```

Chromium does the layout, so Arabic shaping and RTL behave exactly as they do on
screen. `build.py` prints each document twice — once with the running
header/footer and once without — and takes page 1 from the clean pass so the
cover bleeds edge to edge with no furniture on it.

Two Chromium quirks are worth knowing before editing `build.py`: the
header/footer templates are an isolated mini-document that collapses to
`font-size: 0` unless sizes are set in **px** on the elements themselves, and a
large base64 `@font-face` data URI makes the whole template stylesheet fail to
parse — which silently takes the font sizes down with it. The templates
therefore resolve Tajawal by family name from the system font cache;
`ensure_fonts()` installs it from `assets/fonts/` if it is missing.

`build_docx.py` emits native Word text and tables so the document stays
editable, and embeds only the genuinely visual blocks — the Gantt chart and the
colour swatches — as pictures, because Word has no faithful equivalent.

> Word is strict about OOXML child ordering inside `rPr`/`pPr`/`tblPr`/`tcPr`.
> Elements appended at the end produce a file Word tolerates but LibreOffice
> refuses to open. `build_docx.py` inserts at the schema position instead — see
> `_ORDER`. A table cell must also never end on a nested table.

## The logo

Both builds use the real RoboAgentix artwork:

- `assets/logo-light.png` — white on transparent, for the dark cover and the
  Word cover block
- `assets/logo-dark.png` — full-colour on transparent, for the running header

`logo-dark.png` was derived from the supplied white-background artwork by
keying the white out: pixels lighter than luminance 250 become transparent,
anything below 235 stays fully opaque so the brand green keeps its colour, and
only the thin anti-aliased edge band in between gets partial alpha.

If either file is missing the builds fall back to a wordmark drawn in CSS, so
nothing breaks — but the real files are committed and should be used.

## Layout

```
proposals/
  build.py                  # HTML -> PDF
  build_docx.py             # HTML -> Word
  src/01-technical.html     # content
  src/02-commercial.html
  assets/theme.css          # the RoboAgentix document theme
  assets/fonts/             # Tajawal + IBM Plex Sans Arabic
  assets/screens/           # 15 AWNAK screens from the approved prototype
  assets/awnak-mark.png     # AWNAK app mark
  out/                      # generated PDFs and .docx
```

## Editing

Content lives in the two HTML files; nothing is generated from templates, so
prices and dates are edited directly. When changing a figure in the commercial
document, check it in all the places it appears — the headline stat tiles
(§1), the line-item table (§5), the cost distribution table (§5-1), and the
payment plan (§9) all have to agree.

`.keep` wraps a heading and its table so a page break cannot separate them; use
it when a table would otherwise leave one orphaned row on the next page.

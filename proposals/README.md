# RoboAgentix — AWNAK proposal documents

Two Arabic (RTL) proposal documents for **عَوْنَك / AWNAK**, the home-and-family
services platform, produced for RoboAgentix For Software Development.

| Output | Pages | Covers |
| --- | --- | --- |
| `out/RoboAgentix-AWNAK-01-Technical-Proposal-AR.pdf` | 34 | What gets built and with which technology — architecture, stack, data model, dispatch engine, order lifecycle, money layer, integrations, security, performance, delivery artefacts |
| `out/RoboAgentix-AWNAK-02-Commercial-Proposal-AR.pdf` | 21 | Line-item pricing, methodology, schedule, team, payment plan, warranty and support, change control, ownership, acceptance criteria |

Document references are `RAX-TEC-AWNAK-001` and `RAX-FIN-AWNAK-001`, both Rev. A.
The two are written to be read together: every price line in the commercial
document points at a scope described in the technical one.

## Sources

- **Requirements** — the client's 32-section Arabic specification
  (*وثيقة توصيف تطبيق خدمات المنزل والعائلة*). Every section of it is answered
  somewhere in the technical document.
- **Product design** — the approved AWNAK customer-app prototype. Its 15 screens
  were captured into `assets/screens/` and its design tokens (teal `#147D73`,
  Tajawal) are documented in §12 of the technical proposal.
- **Document design** — the RoboAgentix house proposal style: dark forest-green
  cover, black section headings over a green rule, green table headers, stat
  tiles, and a running header/footer carrying the logo, the revision line, the
  company contact line, and a page number.

## Build

```bash
pip install playwright pypdfium2 pillow
python3 proposals/build.py          # both documents
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

## Layout

```
proposals/
  build.py                  # HTML -> PDF
  src/01-technical.html     # content
  src/02-commercial.html
  assets/theme.css          # the RoboAgentix document theme
  assets/fonts/             # Tajawal + IBM Plex Sans Arabic
  assets/screens/           # 15 AWNAK screens from the approved prototype
  assets/awnak-mark.png     # AWNAK app mark
  out/                      # generated PDFs
```

## Editing

Content lives in the two HTML files; nothing is generated from templates, so
prices and dates are edited directly. When changing a figure in the commercial
document, check it in all the places it appears — the headline stat tiles
(§1), the line-item table (§5), the cost distribution table (§5-1), and the
payment plan (§9) all have to agree.

`.keep` wraps a heading and its table so a page break cannot separate them; use
it when a table would otherwise leave one orphaned row on the next page.

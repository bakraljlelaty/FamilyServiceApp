#!/usr/bin/env python3
"""
Render the RoboAgentix proposal HTML sources to print-ready A4 PDFs.

Chromium (bundled with Playwright) does the layout, so Arabic shaping, RTL,
and the Tajawal webfonts behave exactly as they do on screen.

Each document is printed twice — once with the running header/footer and once
without — and page 1 is taken from the clean pass so the cover bleeds edge to
edge with no furniture on it. Page numbers still count the cover, because both
passes are the same document.

  python3 proposals/build.py            # build everything
  python3 proposals/build.py 01         # build one document
"""
import pathlib
import sys

import pypdfium2 as pdfium
from playwright.sync_api import sync_playwright

ROOT = pathlib.Path(__file__).resolve().parent
SRC = ROOT / "src"
OUT = ROOT / "out"
CHROME = "/opt/pw-browsers/chromium-1194/chrome-linux/chrome"
MARGIN = {"top": "26mm", "bottom": "20mm", "left": "16mm", "right": "16mm"}

# Chromium renders the margin boxes as an isolated mini-document. Two traps:
# it collapses to font-size 0 unless sizes are given in px on the elements
# themselves, and a large base64 @font-face data URI makes the whole template
# stylesheet fail to parse (which silently takes the font sizes down with it).
# So the template resolves Tajawal by family name from the system font cache —
# see ensure_fonts(), which installs it if it is missing.
TPL_CSS = """
*{-webkit-print-color-adjust:exact;print-color-adjust:exact;box-sizing:border-box;}
body{margin:0;padding:0;}
table.bar{width:100%;border-collapse:collapse;font-family:'Tajawal',sans-serif;
          font-size:7px;color:#8B968F;}
table.bar td{padding:0;vertical-align:middle;white-space:nowrap;}
.l{text-align:left;} .r{text-align:right;direction:rtl;}
.ra{font-weight:700;font-size:11px;color:#141414;letter-spacing:-.02em;
     white-space:nowrap;direction:ltr;}
.dot{display:inline-block;background:#14512F;border-radius:3.3px;width:15px;height:8.2px;
      position:relative;top:1px;margin:0 .7px;}
.dot i{position:absolute;top:2.7px;width:2.8px;height:2.8px;border-radius:50%;background:#fff;}
.dot i.a{left:3.2px;} .dot i.b{right:3.2px;}
.ix{color:#14512F;}
.conf{color:#14512F;font-weight:700;}
.pg{color:#5F6B65;font-weight:700;direction:ltr;display:inline-block;}
"""

LOGO_FALLBACK = ('<span class="ra">Rob<span class="dot"><i class="a"></i><i class="b"></i></span>'
                 'Agent<span class="ix">ix</span></span>')

# Drop the real artwork in as assets/logo-light.png (white, for the dark cover)
# and assets/logo-dark.png (dark, for the running header) and both the cover and
# the header pick it up automatically. Without them the drawn lockup is used.
LOGO_LIGHT = ROOT / "assets/logo-light.png"
LOGO_DARK = ROOT / "assets/logo-dark.png"


def header_logo() -> str:
    """The running header can only carry the artwork inline, as a data URI.

    Safe here because it goes in the template body — it is specifically a large
    data URI inside <style> that breaks the template stylesheet parse.
    """
    if not LOGO_DARK.exists():
        return LOGO_FALLBACK
    import base64
    b64 = base64.b64encode(LOGO_DARK.read_bytes()).decode()
    return f'<img src="data:image/png;base64,{b64}" style="height:15px;width:auto;display:block;">'


def _bar(left: str, right: str, pad_top: str, pad_bottom: str) -> str:
    return (f'<style>{TPL_CSS}</style>'
            f'<div style="width:100%;padding:{pad_top} 16mm {pad_bottom} 16mm;">'
            f'<table class="bar"><tr>'
            f'<td class="l">{left}</td><td class="r">{right}</td>'
            f'</tr></table></div>')


def header(doc_title: str) -> str:
    return _bar(header_logo(), f'<span>{doc_title}</span>', "11mm", "0")


def footer() -> str:
    right = ('<span class="conf">سرّي تجاريًا</span>'
             '&nbsp;&nbsp;·&nbsp;&nbsp;'
             '<span class="pg"><span class="pageNumber"></span> / '
             '<span class="totalPages"></span></span>')
    left = ('RoboAgentix For Software Development &nbsp;·&nbsp; roboagentix.ai '
            '&nbsp;·&nbsp; contact@roboagentix.ai')
    return _bar(left, right, "0", "9mm")


DOCS = [
    ("01-technical.html", "RoboAgentix-AWNAK-01-Technical-Proposal-AR.pdf",
     "المراجعة (A) — عَوْنَك — العرض الفني التفصيلي"),
    ("02-commercial.html", "RoboAgentix-AWNAK-02-Commercial-Proposal-AR.pdf",
     "المراجعة (A) — عَوْنَك — العرض المالي وآلية العمل والتسليم"),
]


def merge_clean_cover(with_furniture: pathlib.Path, clean: pathlib.Path,
                      target: pathlib.Path) -> int:
    """Page 1 from the clean pass, the rest from the pass with header/footer."""
    a = pdfium.PdfDocument(str(clean))
    b = pdfium.PdfDocument(str(with_furniture))
    out = pdfium.PdfDocument.new()
    out.import_pages(a, [0])
    out.import_pages(b, list(range(1, len(b))))
    n = len(out)
    out.save(str(target))
    for d in (out, a, b):
        d.close()
    return n


def ensure_fonts() -> None:
    """Make Tajawal resolvable by family name, for the header/footer templates.

    The page itself loads the fonts through @font-face from assets/fonts, but
    Chromium's margin-box document cannot, so it needs them in the font cache.
    """
    import shutil
    import subprocess

    if shutil.which("fc-list"):
        installed = subprocess.run(["fc-list", ":family"], capture_output=True, text=True)
        if "Tajawal" in installed.stdout:
            return
    dest = pathlib.Path("/usr/share/fonts/truetype/roboagentix")
    try:
        dest.mkdir(parents=True, exist_ok=True)
        for ttf in (ROOT / "assets/fonts").glob("*.ttf"):
            shutil.copy2(ttf, dest / ttf.name)
        if shutil.which("fc-cache"):
            subprocess.run(["fc-cache", "-f"], capture_output=True)
        print("  · installed Tajawal into the system font cache")
    except OSError as exc:
        print(f"  ! could not install fonts ({exc}); header/footer may fall back")


def build(only: str | None = None) -> None:
    ensure_fonts()
    OUT.mkdir(exist_ok=True)
    tmp = OUT / ".tmp"
    tmp.mkdir(exist_ok=True)
    targets = [d for d in DOCS if only is None or d[0].startswith(only)]
    if not targets:
        sys.exit(f"no source matches {only!r}")

    with sync_playwright() as p:
        browser = p.chromium.launch(executable_path=CHROME, args=["--no-sandbox"])
        page = browser.new_page()
        for src_name, pdf_name, running_title in targets:
            src = SRC / src_name
            if not src.exists():
                print(f"  skip {src_name} (not found)")
                continue
            page.goto(src.as_uri(), wait_until="networkidle", timeout=120_000)
            if LOGO_LIGHT.exists():
                page.evaluate("document.documentElement.classList.add('has-logo')")
            page.emulate_media(media="print")
            page.wait_for_timeout(1500)

            common = dict(format="A4", print_background=True, margin=MARGIN,
                          prefer_css_page_size=False)
            furn = tmp / f"{src_name}.furn.pdf"
            clean = tmp / f"{src_name}.clean.pdf"
            page.pdf(path=str(furn), display_header_footer=True,
                     header_template=header(running_title),
                     footer_template=footer(), **common)
            page.pdf(path=str(clean), display_header_footer=False, **common)

            target = OUT / pdf_name
            pages = merge_clean_cover(furn, clean, target)
            kb = target.stat().st_size / 1024
            print(f"  ✓ {pdf_name}  —  {pages} pages, {kb:,.0f} KB")
        browser.close()

    for f in tmp.glob("*.pdf"):
        f.unlink()
    tmp.rmdir()


if __name__ == "__main__":
    build(sys.argv[1] if len(sys.argv) > 1 else None)

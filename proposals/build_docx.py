#!/usr/bin/env python3
"""
Render the proposal HTML sources to editable Word (.docx) documents.

The PDF is the presentation artefact; this is the working one — prices, dates
and wording stay editable as real Word text and tables. Only the genuinely
visual blocks (the Gantt chart, the architecture layers, the colour swatches)
are embedded as pictures, because Word has no faithful equivalent and
flattening them to text would lose the point.

  python3 proposals/build_docx.py          # both documents
  python3 proposals/build_docx.py 01       # one document
"""
import pathlib
import re
import sys

from bs4 import BeautifulSoup
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from playwright.sync_api import sync_playwright

ROOT = pathlib.Path(__file__).resolve().parent
SRC = ROOT / "src"
OUT = ROOT / "out"
SHOTS = OUT / ".figures"
CHROME = "/opt/pw-browsers/chromium-1194/chrome-linux/chrome"

GREEN = RGBColor(0x14, 0x51, 0x2F)
GREEN_MID = RGBColor(0x1D, 0x6B, 0x47)
INK = RGBColor(0x14, 0x14, 0x14)
INK2 = RGBColor(0x33, 0x3A, 0x36)
INK3 = RGBColor(0x5F, 0x6B, 0x65)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
MIST = RGBColor(0xB5, 0xCC, 0xBE)
PALE = RGBColor(0x8F, 0xBF, 0xA3)
SH_HEAD = "14512F"
SH_TINT = "EEF3F0"
SH_ZEBRA = "F5F8F6"
SH_COVER = "0B2A1C"
SH_AMBER = "FDF6E6"

FONT = "Tajawal"
# Blocks Word cannot reproduce honestly — screenshot these from the rendered page.
FIGURE_SELECTORS = ["table.gantt", "div.sw"]

DOCS = [
    ("01-technical.html", "RoboAgentix-AWNAK-01-Technical-Proposal-AR.docx"),
    ("02-commercial.html", "RoboAgentix-AWNAK-02-Commercial-Proposal-AR.docx"),
]


# ── low-level RTL plumbing ────────────────────────────────────────────────
# OOXML fixes the order of children inside rPr/pPr/tblPr/tcPr. Appending at the
# end produces a file Word tolerates but LibreOffice refuses to open outright,
# so every element has to be inserted at its schema position.
_ORDER = {
    "w:rPr": ["w:rStyle", "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:caps",
              "w:smallCaps", "w:strike", "w:dstrike", "w:outline", "w:shadow",
              "w:emboss", "w:imprint", "w:noProof", "w:snapToGrid", "w:vanish",
              "w:webHidden", "w:color", "w:spacing", "w:w", "w:kern", "w:position",
              "w:sz", "w:szCs", "w:highlight", "w:u", "w:effect", "w:bdr", "w:shd",
              "w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang"],
    "w:pPr": ["w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore",
              "w:framePr", "w:widowControl", "w:numPr", "w:suppressLineNumbers",
              "w:pBdr", "w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku",
              "w:wordWrap", "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE",
              "w:autoSpaceDN", "w:bidi", "w:adjustRightInd", "w:snapToGrid",
              "w:spacing", "w:ind", "w:contextualSpacing", "w:mirrorIndents",
              "w:suppressOverlap", "w:jc", "w:textDirection", "w:textAlignment",
              "w:outlineLvl", "w:rPr", "w:sectPr"],
    "w:tblPr": ["w:tblStyle", "w:tblpPr", "w:tblOverlap", "w:bidiVisual",
                "w:tblStyleRowBandSize", "w:tblStyleColBandSize", "w:tblW", "w:jc",
                "w:tblCellSpacing", "w:tblInd", "w:tblBorders", "w:shd",
                "w:tblLayout", "w:tblCellMar", "w:tblLook"],
    "w:tcPr": ["w:cnfStyle", "w:tcW", "w:gridSpan", "w:hMerge", "w:vMerge",
               "w:tcBorders", "w:shd", "w:noWrap", "w:tcMar", "w:textDirection",
               "w:tcFitText", "w:vAlign", "w:hideMark"],
}


def _set(el, tag, **attrs):
    """Insert (or update) a child element at its schema-mandated position."""
    node = el.find(qn(tag))
    if node is None:
        node = OxmlElement(tag)
        order = _ORDER.get(el.tag.split("}")[-1] and f"w:{el.tag.split('}')[-1]}")
        if order and tag in order:
            rank = order.index(tag)
            anchor = None
            for child in el:
                name = f"w:{child.tag.split('}')[-1]}"
                if name in order and order.index(name) > rank:
                    anchor = child
                    break
            if anchor is not None:
                anchor.addprevious(node)
            else:
                el.append(node)
        else:
            el.append(node)
    for k, v in attrs.items():
        node.set(qn(k), v)
    return node


def rtl_para(p):
    pPr = p._p.get_or_add_pPr()
    _set(pPr, "w:bidi", **{"w:val": "1"})
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    return p


def style_run(r, size=10.5, bold=False, color=INK2, font=FONT):
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    rPr = r._r.get_or_add_rPr()
    rf = rPr.find(qn("w:rFonts"))
    if rf is None:
        rf = _set(rPr, "w:rFonts")
    rf.set(qn("w:cs"), font)
    rf.set(qn("w:ascii"), font)
    rf.set(qn("w:hAnsi"), font)
    _set(rPr, "w:rtl", **{"w:val": "1"})
    _set(rPr, "w:szCs", **{"w:val": str(int(size * 2))})
    return r


def shade(cell, hexcolor):
    _set(cell._tc.get_or_add_tcPr(), "w:shd",
         **{"w:val": "clear", "w:color": "auto", "w:fill": hexcolor})


def rtl_table(t):
    _set(t._tbl.tblPr, "w:bidiVisual", **{"w:val": "1"})
    t.alignment = WD_TABLE_ALIGNMENT.RIGHT
    return t


def cell_text(cell, text, size=9, bold=False, color=INK2, align=None):
    cell.text = ""
    p = rtl_para(cell.paragraphs[0])
    if align:
        p.alignment = align
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    style_run(p.add_run(text), size=size, bold=bold, color=color)
    return p


def bottom_border(p, hexcolor=SH_HEAD, size="12"):
    pPr = p._p.get_or_add_pPr()
    borders = _set(pPr, "w:pBdr")
    _set(borders, "w:bottom", **{"w:val": "single", "w:sz": size,
                                 "w:space": "3", "w:color": hexcolor})


# ── text helpers ──────────────────────────────────────────────────────────
def clean(s):
    return re.sub(r"\s+", " ", s or "").strip()


def add_para(doc, node, size=10.5, color=INK2, space_after=5, bold_all=False):
    """Emit one paragraph, keeping <strong>/<em>/<code> emphasis as real runs."""
    p = rtl_para(doc.add_paragraph())
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.45
    empty = True
    for part in node.children if hasattr(node, "children") else []:
        if getattr(part, "name", None) is None:
            txt = clean(str(part))
            if txt:
                style_run(p.add_run(txt + " "), size=size, color=color, bold=bold_all)
                empty = False
        else:
            txt = clean(part.get_text())
            if not txt:
                continue
            name = part.name
            cls = part.get("class") or []
            bold = bold_all or name in ("strong", "b") or "t" in cls
            col = GREEN_MID if (name in ("em", "code") or "gr" in cls) else color
            style_run(p.add_run(txt + " "), size=size, bold=bold, color=col)
            empty = False
    if empty:
        txt = clean(node.get_text() if hasattr(node, "get_text") else str(node))
        if txt:
            style_run(p.add_run(txt), size=size, color=color, bold=bold_all)
    return p


def add_list(doc, ul):
    for li in ul.find_all("li", recursive=False):
        p = rtl_para(doc.add_paragraph())
        p.paragraph_format.space_after = Pt(1.5)
        p.paragraph_format.left_indent = Inches(0.12)
        p.paragraph_format.right_indent = Inches(0.16)
        p.paragraph_format.line_spacing = 1.35
        style_run(p.add_run("•  "), size=9.5, color=GREEN_MID, bold=True)
        for part in li.children:
            if getattr(part, "name", None) is None:
                t = clean(str(part))
                if t:
                    style_run(p.add_run(t + " "), size=9.5)
            else:
                t = clean(part.get_text())
                if not t:
                    continue
                bold = part.name in ("strong", "b")
                col = GREEN_MID if part.name in ("em", "code") else INK2
                style_run(p.add_run(t + " "), size=9.5, bold=bold, color=col)


def add_table(doc, tbl):
    head = tbl.find("thead")
    heads = [clean(th.get_text()) for th in head.find_all("th")] if head else []
    body = tbl.find("tbody") or tbl
    rows = [r for r in body.find_all("tr") if r.find(["td", "th"])]
    if not rows:
        return
    ncols = max(len(r.find_all(["td", "th"])) for r in rows) or len(heads)
    ncols = max(ncols, len(heads))
    t = doc.add_table(rows=0, cols=ncols)
    t.style = "Table Grid"
    rtl_table(t)

    if heads:
        cells = t.add_row().cells
        for i, h in enumerate(heads[:ncols]):
            shade(cells[i], SH_HEAD)
            cell_text(cells[i], h, size=8.5, bold=True, color=WHITE)

    for n, tr in enumerate(rows):
        tds = tr.find_all(["td", "th"])
        cells = t.add_row().cells
        cls = tr.get("class") or []
        total = "tot" in cls or "sub" in cls
        fill = SH_TINT if total else (SH_ZEBRA if n % 2 else None)
        ci = 0
        for td in tds:
            if ci >= ncols:
                break
            span = int(td.get("colspan", 1))
            span = min(span, ncols - ci)
            target = cells[ci]
            if span > 1:
                target = target.merge(cells[ci + span - 1])
            tdcls = td.get("class") or []
            align = (WD_ALIGN_PARAGRAPH.CENTER if "mid" in tdcls else
                     WD_ALIGN_PARAGRAPH.LEFT if "num" in tdcls else None)
            cell_text(target, clean(td.get_text()), size=8.5, bold=total,
                      color=GREEN if total else INK2, align=align)
            if fill:
                for k in range(ci, ci + span):
                    shade(cells[k], fill)
            ci += span
    doc.add_paragraph().paragraph_format.space_after = Pt(3)


def boxed(doc, text_nodes, fill=SH_TINT, title=None):
    """A note/callout — one shaded, full-width cell."""
    t = doc.add_table(rows=1, cols=1)
    t.style = "Table Grid"
    rtl_table(t)
    c = t.rows[0].cells[0]
    shade(c, fill)
    c.text = ""
    first = True
    for kind, node in text_nodes:
        p = c.paragraphs[0] if first else c.add_paragraph()
        rtl_para(p)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.4
        first = False
        if kind == "title":
            style_run(p.add_run(clean(node)), size=9.5, bold=True, color=GREEN)
        else:
            for part in node.children:
                if getattr(part, "name", None) is None:
                    tx = clean(str(part))
                    if tx:
                        style_run(p.add_run(tx + " "), size=9)
                else:
                    tx = clean(part.get_text())
                    if not tx or (part.get("class") or []) == ["t"]:
                        continue
                    style_run(p.add_run(tx + " "), size=9,
                              bold=part.name in ("strong", "b"),
                              color=GREEN_MID if part.name in ("em", "code") else INK2)
    doc.add_paragraph().paragraph_format.space_after = Pt(3)


def grid_to_table(doc, container, item_sel, title_sel, body_sel, per_row=None):
    items = container.select(item_sel)
    if not items:
        return
    per_row = per_row or min(len(items), 3)
    rows = (len(items) + per_row - 1) // per_row
    t = doc.add_table(rows=0, cols=per_row)
    t.style = "Table Grid"
    rtl_table(t)
    for r in range(rows):
        cells = t.add_row().cells
        for c in range(per_row):
            i = r * per_row + c
            if i >= len(items):
                cell_text(cells[c], "")
                continue
            it = items[i]
            head = it.select_one(title_sel)
            bodyn = it.select_one(body_sel)
            cells[c].text = ""
            if head:
                p = rtl_para(cells[c].paragraphs[0])
                p.paragraph_format.space_after = Pt(1)
                style_run(p.add_run(clean(head.get_text())), size=9, bold=True, color=GREEN)
            if bodyn:
                lis = bodyn.find_all("li")
                if lis:
                    for li in lis:
                        p = rtl_para(cells[c].add_paragraph())
                        p.paragraph_format.space_after = Pt(0.5)
                        style_run(p.add_run("• " + clean(li.get_text())), size=8.5)
                else:
                    p = rtl_para(cells[c].add_paragraph())
                    p.paragraph_format.space_after = Pt(1)
                    style_run(p.add_run(clean(bodyn.get_text())), size=8.5)
            shade(cells[c], SH_ZEBRA)
    doc.add_paragraph().paragraph_format.space_after = Pt(3)


# ── figures ───────────────────────────────────────────────────────────────
def capture_figures(html_path, tag):
    SHOTS.mkdir(parents=True, exist_ok=True)
    shots = {}
    with sync_playwright() as pw:
        b = pw.chromium.launch(executable_path=CHROME, args=["--no-sandbox"])
        pg = b.new_page(viewport={"width": 1200, "height": 1400}, device_scale_factor=3)
        pg.goto(html_path.as_uri(), wait_until="networkidle", timeout=120_000)
        pg.wait_for_timeout(1200)
        for sel in FIGURE_SELECTORS:
            for i, el in enumerate(pg.query_selector_all(sel)):
                key = f"{sel}#{i}"
                out = SHOTS / f"{tag}-{sel.replace('.','_').replace(' ','')}-{i}.png"
                try:
                    el.screenshot(path=str(out))
                    shots[key] = out
                except Exception:
                    pass
        b.close()
    return shots


# ── document assembly ─────────────────────────────────────────────────────
def build_cover(doc, cover):
    t = doc.add_table(rows=1, cols=1)
    rtl_table(t)
    c = t.rows[0].cells[0]
    shade(c, SH_COVER)
    c.text = ""

    def line(text, size, bold, color, after=6, first=False):
        p = c.paragraphs[0] if first else c.add_paragraph()
        rtl_para(p)
        p.paragraph_format.space_after = Pt(after)
        style_run(p.add_run(text), size=size, bold=bold, color=color)

    line("RoboAgentix", 20, True, WHITE, after=26, first=True)
    badge = cover.select_one(".badge")
    if badge:
        line(clean(badge.get_text()), 9, False, PALE, after=4)
    line(clean(cover.select_one("h1").get_text()), 26, True, WHITE, after=6)
    sub = cover.select_one(".sub")
    if sub:
        line(clean(sub.get_text()), 14, False, PALE, after=8)
    desc = cover.select_one(".desc")
    if desc:
        p = c.add_paragraph()
        rtl_para(p)
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.space_after = Pt(14)
        style_run(p.add_run(clean(desc.get_text())), size=10, color=MIST)

    meta = cover.select(".meta > div")
    if meta:
        mt = c.add_table(rows=2, cols=len(meta))
        rtl_table(mt)
        for i, m in enumerate(meta):
            k = m.select_one(".k")
            v = m.select_one(".v")
            cell_text(mt.rows[0].cells[i], clean(k.get_text()) if k else "",
                      size=8, color=PALE)
            cell_text(mt.rows[1].cells[i], clean(v.get_text()) if v else "",
                      size=9.5, bold=True, color=WHITE)
            shade(mt.rows[0].cells[i], SH_COVER)
            shade(mt.rows[1].cells[i], SH_COVER)

    # A table cell must not end on a nested table — always close with a paragraph.
    foot = cover.select_one(".cfoot")
    p = c.add_paragraph()
    rtl_para(p)
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.line_spacing = 1.4
    if foot:
        style_run(p.add_run(clean(foot.get_text())), size=8, color=MIST)
    doc.add_page_break()


def convert(html_path, docx_path, tag):
    soup = BeautifulSoup(html_path.read_text(encoding="utf-8"), "html.parser")
    figures = capture_figures(html_path, tag)

    doc = Document()
    sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Inches(0.75)
    sec.top_margin = sec.bottom_margin = Inches(0.8)
    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(10.5)

    cover = soup.select_one(".cover")
    if cover:
        build_cover(doc, cover)

    fig_counts = {s: 0 for s in FIGURE_SELECTORS}
    for section in soup.find_all("section"):
        for node in section.children:
            if getattr(node, "name", None) is None:
                continue
            cls = node.get("class") or []
            name = node.name

            if name == "h2":
                p = rtl_para(doc.add_paragraph())
                p.paragraph_format.space_before = Pt(16)
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.keep_with_next = True
                style_run(p.add_run(clean(node.get_text())), size=18, bold=True, color=INK)
                bottom_border(p)
            elif name == "h3":
                p = rtl_para(doc.add_paragraph())
                p.paragraph_format.space_before = Pt(11)
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.keep_with_next = True
                style_run(p.add_run(clean(node.get_text())), size=12, bold=True, color=GREEN_MID)
            elif name == "h4":
                p = rtl_para(doc.add_paragraph())
                p.paragraph_format.space_before = Pt(8)
                p.paragraph_format.space_after = Pt(2)
                style_run(p.add_run(clean(node.get_text())), size=10.5, bold=True, color=INK)
            elif name == "p":
                if "lead" in cls:
                    add_para(doc, node, size=11, color=INK3, space_after=7)
                elif "tnote" in cls or "figcap" in cls or "xs" in cls:
                    add_para(doc, node, size=8.5, color=INK3, space_after=7)
                elif "small" in cls:
                    add_para(doc, node, size=9, space_after=5)
                else:
                    add_para(doc, node)
            elif name in ("ul", "ol"):
                add_list(doc, node)
            elif name == "table":
                if "gantt" in cls:
                    i = fig_counts["table.gantt"]
                    fig_counts["table.gantt"] += 1
                    img = figures.get(f"table.gantt#{i}")
                    if img and img.exists():
                        doc.add_picture(str(img), width=Inches(6.6))
                        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    continue
                add_table(doc, node)
            elif "note" in cls:
                title = node.select_one(".t")
                boxed(doc, ([("title", title.get_text())] if title else []) + [("body", node)],
                      fill=SH_AMBER if "warn" in cls else SH_TINT)
            elif "stats" in cls:
                items = node.select(".s")
                t = doc.add_table(rows=2, cols=len(items))
                t.style = "Table Grid"
                rtl_table(t)
                for i, it in enumerate(items):
                    v = it.select_one(".v")
                    l = it.select_one(".l")
                    cell_text(t.rows[0].cells[i], clean(v.get_text()) if v else "",
                              size=15, bold=True, color=GREEN,
                              align=WD_ALIGN_PARAGRAPH.CENTER)
                    cell_text(t.rows[1].cells[i], clean(l.get_text()) if l else "",
                              size=8, color=INK3, align=WD_ALIGN_PARAGRAPH.CENTER)
                    shade(t.rows[0].cells[i], SH_TINT)
                    shade(t.rows[1].cells[i], SH_TINT)
                doc.add_paragraph().paragraph_format.space_after = Pt(3)
            elif "cards" in cls:
                per = 3 if "c3" in cls else 2
                grid_to_table(doc, node, ".card", ".h", ".b", per_row=per)
            elif "steps" in cls:
                t = doc.add_table(rows=0, cols=2)
                t.style = "Table Grid"
                rtl_table(t)
                for st in node.select(".step"):
                    n = st.select_one(".n")
                    ti = st.select_one(".t")
                    d = st.select_one(".d")
                    cells = t.add_row().cells
                    cell_text(cells[0], clean(n.get_text()) if n else "", size=9.5,
                              bold=True, color=GREEN, align=WD_ALIGN_PARAGRAPH.CENTER)
                    cells[0].width = Inches(0.4)
                    cells[1].text = ""
                    p = rtl_para(cells[1].paragraphs[0])
                    p.paragraph_format.space_after = Pt(1)
                    style_run(p.add_run(clean(ti.get_text()) if ti else ""),
                              size=9.5, bold=True, color=INK)
                    if d:
                        p2 = rtl_para(cells[1].add_paragraph())
                        p2.paragraph_format.space_after = Pt(1)
                        style_run(p2.add_run(clean(d.get_text())), size=9, color=INK3)
                doc.add_paragraph().paragraph_format.space_after = Pt(3)
            elif "layer" in cls:
                head = node.select_one(".lh .a")
                boxes = node.select(".box")
                if head:
                    p = rtl_para(doc.add_paragraph())
                    p.paragraph_format.space_before = Pt(4)
                    p.paragraph_format.space_after = Pt(1)
                    style_run(p.add_run(clean(head.get_text())), size=9.5, bold=True, color=GREEN)
                if boxes:
                    p = rtl_para(doc.add_paragraph())
                    p.paragraph_format.space_after = Pt(4)
                    p.paragraph_format.right_indent = Inches(0.16)
                    style_run(p.add_run(" · ".join(
                        clean(b.find("b").get_text()) for b in boxes if b.find("b"))), size=9)
            elif "sw" in cls:
                i = fig_counts["div.sw"]
                fig_counts["div.sw"] += 1
                img = figures.get(f"div.sw#{i}")
                if img and img.exists():
                    doc.add_picture(str(img), width=Inches(6.4))
                    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif "shots" in cls:
                shots = node.select(".shot")
                t = doc.add_table(rows=2, cols=len(shots))
                rtl_table(t)
                for i, sh in enumerate(shots):
                    img = sh.find("img")
                    cap = sh.select_one(".cap")
                    cells0 = t.rows[0].cells[i]
                    cells0.text = ""
                    if img and img.get("src"):
                        pth = (html_path.parent / img["src"]).resolve()
                        if pth.exists():
                            run = rtl_para(cells0.paragraphs[0]).add_run()
                            run.add_picture(str(pth), width=Inches(1.4))
                    cell_text(t.rows[1].cells[i], clean(cap.get_text()) if cap else "",
                              size=7.5, color=INK3, align=WD_ALIGN_PARAGRAPH.CENTER)
                doc.add_paragraph().paragraph_format.space_after = Pt(3)
            elif "states" in cls or "pills" in cls:
                chips = node.select(".st") or node.select(".pill")
                p = rtl_para(doc.add_paragraph())
                p.paragraph_format.space_after = Pt(6)
                p.paragraph_format.line_spacing = 1.5
                style_run(p.add_run("  ·  ".join(clean(c.get_text()) for c in chips)),
                          size=9, color=GREEN_MID)
            elif "toc" in cls:
                for item in node.select(".i"):
                    n = item.select_one(".n")
                    tt = item.select_one(".t")
                    p = rtl_para(doc.add_paragraph())
                    p.paragraph_format.space_after = Pt(1.5)
                    style_run(p.add_run((clean(n.get_text()) + "   ") if n else ""),
                              size=9.5, bold=True, color=GREEN_MID)
                    style_run(p.add_run(clean(tt.get_text()) if tt else ""), size=9.5)
            elif "sign" in cls:
                blocks = node.select(".s")
                t = doc.add_table(rows=1, cols=len(blocks))
                t.style = "Table Grid"
                rtl_table(t)
                for i, b in enumerate(blocks):
                    r = b.select_one(".r")
                    l = b.select_one(".l")
                    c = t.rows[0].cells[i]
                    c.text = ""
                    p = rtl_para(c.paragraphs[0])
                    style_run(p.add_run(clean(r.get_text()) if r else ""),
                              size=9, bold=True, color=INK)
                    c.add_paragraph()
                    c.add_paragraph()
                    p3 = rtl_para(c.add_paragraph())
                    style_run(p3.add_run(clean(l.get_text()) if l else ""),
                              size=8, color=INK3)
            elif "keep" in cls:
                for sub in node.children:
                    if getattr(sub, "name", None) == "h3":
                        p = rtl_para(doc.add_paragraph())
                        p.paragraph_format.space_before = Pt(11)
                        p.paragraph_format.keep_with_next = True
                        style_run(p.add_run(clean(sub.get_text())), size=12,
                                  bold=True, color=GREEN_MID)
                    elif getattr(sub, "name", None) == "table":
                        add_table(doc, sub)
            elif "endmark" in cls:
                p = rtl_para(doc.add_paragraph())
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(16)
                style_run(p.add_run(clean(node.get_text())), size=8.5, color=INK3)
            elif name == "div" and "h2rule" in cls:
                continue

        if "page-break" in (section.get("class") or []):
            pass

    doc.save(str(docx_path))
    return docx_path


def main(only=None):
    OUT.mkdir(exist_ok=True)
    targets = [d for d in DOCS if only is None or d[0].startswith(only)]
    if not targets:
        sys.exit(f"no source matches {only!r}")
    for src_name, docx_name in targets:
        src = SRC / src_name
        if not src.exists():
            print(f"  skip {src_name} (not found)")
            continue
        out = convert(src, OUT / docx_name, src_name[:2])
        print(f"  ✓ {docx_name}  —  {out.stat().st_size/1024:,.0f} KB")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else None)

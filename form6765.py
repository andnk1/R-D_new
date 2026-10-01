"""
form6765.py - draft Form 6765 (flattened PDF) from the study's engine result.

- Uses the official IRS fillable form in forms/f6765.pdf (Rev. December 2024). The field
  positions are read from that file, the values are drawn onto the page, and the fillable
  fields are removed (flattened).
- Only the method used in the study is completed (Section A = Regular, Section B = ASC).
- Whole dollars. Lines the app does not collect are left blank for the preparer.
- Red notes and a "Draft – review before filing" stamp are added for the preparer.

When the IRS issues a new revision: replace forms/f6765.pdf and check FIELDS below
(the field names are listed by running:  python form6765.py --fields).
"""
import io
import os
import sys
from datetime import datetime

from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject
from reportlab.lib.colors import Color
from reportlab.pdfgen import canvas

import study as S
from engine import rd_feasibility_generator as gen

HERE = os.path.dirname(os.path.abspath(__file__))
FORM_PATH = os.path.join(HERE, "forms", "f6765.pdf")
FORM_REV = "Rev. December 2024"

RED = Color(0.70, 0, 0)
BLUE = Color(0, 0.13, 0.40)

# Form line -> field name on the IRS PDF (page 1 = f1_, page 2 = f2_)
FIELDS = {
    "name": "f1_01", "ein": "f1_02",
    **{str(n): f"f1_{n + 2:02d}" for n in range(1, 27)},          # lines 1-26 -> f1_03..f1_28
    "27": "f2_01", "28": "f2_02", "29": "f2_03", "30": "f2_04", "31": "f2_05", "32": "f2_06",
    "34": "f2_07", "35": "f2_08", "36": "f2_09", "37": "f2_10", "38": "f2_11", "41amt": "f2_12",
    "42": "f2_13", "43": "f2_14", "44": "f2_15", "45": "f2_16", "46": "f2_17", "47": "f2_18", "48": "f2_19",
}
# Check boxes: (field, index) - index 0 = "Yes" box, 1 = "No" box
CHECKS = {
    "A_yes": ("c1_1", 0), "A_no": ("c1_1", 1),
    "B_yes": ("c1_2", 0), "B_no": ("c1_2", 1),
    "33a": ("c2_1", 0), "33b": ("c2_2", 0),
    "FA_yes": ("c2_6", 0), "FA_no": ("c2_6", 1),
}

SMALL_FILER_QRE = 1_500_000
SMALL_FILER_GR = 50_000_000
PAYROLL_MAX = 500_000


def r(v):
    """Whole dollars (half up)."""
    return None if v is None else int(v + 0.5) if v >= 0 else -int(-v + 0.5)


# ---------------------------------------------------------------------------
# Line values
# ---------------------------------------------------------------------------

def build(study):
    """Return dict: lines {line: value}, checks [names], notes [(where, text)], meta."""
    data, res, used = S.calculate(study)
    c = study["company"]
    ty = S.tax_year(study)
    L, checks, notes, sheet = {}, [], [], []          # sheet = preparer-notes box lines

    L["name"] = c.get("name") or ""
    L["ein"] = c.get("ein") or ""

    # Item A - the engine always uses the reduced (280C) credit
    checks.append("A_yes")
    # Item B - controlled group
    cg = c.get("cg", "no")
    if cg == "yes":
        checks.append("B_yes")
    elif cg == "no":
        checks.append("B_no")
    if cg in ("yes", "unsure"):
        notes.append(("itemB", "Controlled group: “%s” – see notes" % cg.capitalize()))
        sheet.append("Item B: controlled group answered “%s” on step 1. Group calculations and the required Item B "
                     "attachment are not prepared by the app – review before filing." % cg.capitalize())

    # Section F (QRE summary) – whole dollars, line 46 (basic research) not collected
    L["42"] = r(res["qre_wages"])
    L["43"] = r(res["qre_supplies"])
    L["44"] = r(res["qre_computers"])
    L["45"] = r(res["qualified_contract_US"])
    L["47"] = L["45"]
    L["48"] = L["42"] + L["43"] + L["44"] + L["47"]

    # Section G required?  Small-filer thresholds: QREs <= $1.5M AND average GR (prior 3 yrs) <= $50M
    gr3 = [data.get(f"annual_gross_receipts_yr_minus{i}") for i in (1, 2, 3)]
    avg3 = (sum(gr3) / 3) if all(v is not None for v in gr3) else None
    small = L["48"] <= SMALL_FILER_QRE and avg3 is not None and avg3 <= SMALL_FILER_GR
    if small or ty < 2026:
        checks.append("FA_no")
    g_note = ("Section G not completed. Section G is not required for small filers: total QREs (line 48) of "
              "$1.5 million or less AND average annual gross receipts of $50 million or less for the prior three "
              "tax years. It is also optional for tax years beginning before 2026.")
    if not small:
        why = "average gross receipts for the prior 3 years were not all entered" if avg3 is None else \
              "this company is above the small-filer thresholds"
        g_note += f" Check: {why}" + (" – Section G is required for tax years beginning after 2025." if ty >= 2026 else ".")
    notes.append(("secF", "Section G not completed – see note on page 3."))
    sheet.append(g_note)

    # Section A or B – only the method used in the study
    if used == "regular":
        L["5"] = L["48"]
        pct = res["regular_fixed_base_pct"]
        L["6"] = f"{pct * 100:.2f}"
        L["7"] = r(res["regular_avg_gross_receipts"])
        L["8"] = r(L["7"] * pct)
        L["9"] = max(L["5"] - L["8"], 0)
        L["10"] = r(L["5"] * 0.50)
        L["11"] = min(L["9"], L["10"])
        L["12"] = L["11"]                      # lines 1 and 4 not collected (blank)
        L["13"] = r(L["12"] * 0.158)
        credit = L["13"]
        notes.append(("secB", "Not completed – Regular credit (Section A) used."))
        if (res.get("qre_year_number") or 0) >= 11:
            sheet.append("Line 6: QRE year 11 or later – the engine uses QRE years 5–9 for the fixed-base "
                         "percentage; the rule allows any 5 of years 5–10. Confirm before filing.")
    elif used == "asc":
        L["20"] = L["48"]
        L["21"] = r(res["asc_prior3_total"])
        if res["asc_startup"]:
            L["24"] = r(L["20"] * 0.06)        # no QREs in one of the prior 3 years: skip 22-23
        else:
            L["22"] = r(L["21"] / 6.0)
            L["23"] = max(L["20"] - L["22"], 0)
            L["24"] = r(L["23"] * 0.14)
        L["25"] = L["24"]                      # lines 14-19 not collected (blank)
        L["26"] = r(L["25"] * 0.79)
        credit = L["26"]
        notes.append(("secA", "Not completed – ASC (Section B) used."))
    else:
        raise ValueError("No credit could be calculated – complete the study inputs first.")

    # Section C – line 27 (Form 8932) and line 29 (pass-through credit) left blank
    L["28"] = credit
    L["30"] = credit

    # Section D – payroll tax election
    entity = c.get("entity", "C")
    gr = {"current": data.get("annual_gross_receipts_current"), "return_filed": c.get("filed", "no")}
    for i in range(1, 16):
        gr[f"yr_minus{i}"] = data.get(f"annual_gross_receipts_yr_minus{i}")
    eligible = gen.qsb_eligible(gr) and c.get("filed", "no") == "no"
    show = study.get("other", {}).get("show_payroll", "yes") == "yes"
    prep_msg = "Tax preparer should complete this part if payroll tax is to be taken."
    if show and eligible:
        checks.append("33a")
        L["34"] = min(credit, PAYROLL_MAX)
        if entity in ("S", "P"):
            L["36"] = min(L["28"], L["34"])
            sheet.append("Section D: line 34 shows the largest amount allowed (the smaller of line 28 or $500,000). "
                         "Enter a smaller amount if the client elects less.")
        else:
            notes.append(("line35", "35–36: complete after Form 3800"))
            sheet.append("Section D: line 34 shows the largest amount allowed (the smaller of line 28 or $500,000). "
                         "Lines 35 and 36 need the general business credit carryforward from Form 3800 – complete them after Form 3800.")
        sheet.append("Section D: check line 33b if payroll tax is reported under a different EIN. "
                     "Prior payroll elections (5-year limit) and controlled groups are not checked by the app.")
    elif show and not eligible:
        notes.append(("secD", prep_msg))
        sheet.append("Section D left blank: the company does not appear to meet the qualified small business tests "
                     "(current-year gross receipts under $5 million and no gross receipts before the 5-year period) "
                     "or the return has already been filed. " + prep_msg)
    else:
        notes.append(("secD", prep_msg))
        sheet.append("Section D left blank (step 5: payroll option not shown). " + prep_msg)

    # Section E – only line 38 is collected
    ow = S.officer_wages(study)
    if ow is not None:
        L["38"] = r(ow)
    else:
        sheet.append("Line 38: wages were entered as a total only, so officers' wages are unknown.")
    notes.append(("secE", "Review questions 37–41, no data collected."))

    sheet.append("Left blank (not collected): lines 1–4 and 14–19 (energy consortia / basic research), "
                 "27 (Form 8932), 29 (pass-through credit), 46 (basic research payments).")
    if res["recommended_note"].endswith("(other method gives a higher credit)"):
        sheet.append("Method: the preparer chose the method that gives the LOWER credit.")

    return {"lines": L, "checks": checks, "notes": notes, "sheet": sheet, "method": used,
            "company": L["name"] or "Company", "tax_year": ty}


# ---------------------------------------------------------------------------
# PDF
# ---------------------------------------------------------------------------

def _widgets(reader):
    """{(field name, index): (page no, rect)} read from the IRS form."""
    out, seen = {}, {}
    for pi, page in enumerate(reader.pages):
        for a in page.get("/Annots", []) or []:
            a = a.get_object()
            if a.get("/Subtype") != "/Widget":
                continue
            t = a.get("/T") or a.get("/Parent", {}).get("/T")
            base = str(t).split("[")[0]
            idx = seen.get(base, 0)
            seen[base] = idx + 1
            out[(base, idx)] = (pi, [float(x) for x in a["/Rect"]])
    return out


def _fmt(v):
    if isinstance(v, (int, float)):
        return f"{v:,.0f}"
    return str(v)


def _wrap(cv, text, font, size, width):
    words, lines, cur = text.split(), [], ""
    for w in words:
        t = (cur + " " + w).strip()
        if cv.stringWidth(t, font, size) <= width:
            cur = t
        else:
            lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines


# where the short red notes go: (page, x, y, width)
NOTE_POS = {
    "itemB": (0, 420, 651, 156),
    "secA":  (0, 300, 628, 200),
    "secB":  (0, 415, 429, 160),
    "secD":  (1, 507, 462, 68, 5.6),
    "line35": (1, 507, 404, 68),
    "secE":  (1, 240, 333, 336),
    "secF":  (1, 330, 189, 246),
}


def render(info):
    reader = PdfReader(FORM_PATH)
    w = _widgets(reader)
    n = len(reader.pages)
    W, H = [float(p.mediabox.width) for p in reader.pages], [float(p.mediabox.height) for p in reader.pages]

    buf = io.BytesIO()
    cv = canvas.Canvas(buf, pagesize=(W[0], H[0]))
    for pi in range(n):
        cv.setPageSize((W[pi], H[pi]))
        # Draft stamp
        cv.setFillColor(RED)
        cv.setFont("Helvetica-Bold", 10)
        cv.drawCentredString(W[pi] / 2, H[pi] - 22, "DRAFT – review before filing")
        cv.setFont("Helvetica", 6.5)
        cv.drawCentredString(W[pi] / 2, H[pi] - 31,
                             f"Prepared from the R&D feasibility study · {info['company']} · tax year {info['tax_year']} · "
                             f"{datetime.now().strftime('%b %d, %Y')} · Form 6765 ({FORM_REV})")
        # Values
        cv.setFillColor(BLUE)
        for key, val in info["lines"].items():
            if val is None or val == "":
                continue
            fld = FIELDS[key]
            p, (x1, y1, x2, y2) = w[(fld, 0)]
            if p != pi:
                continue
            txt = _fmt(val)
            size = 9
            cv.setFont("Helvetica", size)
            if key in ("name",):
                cv.drawString(x1 + 2, y1 + 3, txt)
            elif key == "ein":
                cv.drawCentredString((x1 + x2) / 2, y1 + 3, txt)
            else:
                cv.drawRightString(x2 - 3, y1 + 3, txt)
        for ck in info["checks"]:
            fld, idx = CHECKS[ck]
            p, (x1, y1, x2, y2) = w[(fld, idx)]
            if p != pi:
                continue
            cv.setFont("Helvetica-Bold", 9)
            cv.drawCentredString((x1 + x2) / 2, y1 + 1.2, "X")
        # Short red notes
        cv.setFillColor(RED)
        for where, text in info["notes"]:
            p, x, y, width, *fs = NOTE_POS[where]
            if p != pi:
                continue
            size = fs[0] if fs else 6.5
            cv.setFont("Helvetica-Bold", size)
            lines = _wrap(cv, text, "Helvetica-Bold", size, width)
            for k, ln in enumerate(lines):
                cv.drawString(x, y - k * size * 1.15, ln)
        # Preparer notes box – bottom of page 1 (space below the form)
        if pi == 0:
            _notes_box(cv, 36, 205, 540, "Preparer notes (not part of the form)", info["sheet"])
        if pi in (2, 3):
            top = 232 if pi == 2 else 184
            _notes_box(cv, 36, top, 540, "Section G – not completed",
                       ["Section G is not required for small filers: total QREs of $1.5 million or less AND average "
                        "annual gross receipts of $50 million or less for the prior three tax years (and it is optional for "
                        "tax years beginning before 2026). See the preparer notes on page 1."])
        cv.showPage()
    cv.save()
    buf.seek(0)
    overlay = PdfReader(buf)

    out = PdfWriter()
    for pi, page in enumerate(reader.pages):
        if "/Annots" in page:                       # flatten: drop the fillable fields
            del page[NameObject("/Annots")]
        page.merge_page(overlay.pages[pi])
        out.add_page(page)
    if "/AcroForm" in out._root_object:
        del out._root_object[NameObject("/AcroForm")]
    out.add_metadata({"/Title": f"Form 6765 DRAFT - {info['company']} {info['tax_year']}"})
    b = io.BytesIO()
    out.write(b)
    return b.getvalue()


def _notes_box(cv, x, top, width, title, items):
    cv.setFillColor(RED)
    cv.setFont("Helvetica-Bold", 7.5)
    cv.drawString(x + 6, top - 10, title)
    y = top - 20
    cv.setFont("Helvetica", 6.8)
    for it in items:
        lines = _wrap(cv, it, "Helvetica", 6.8, width - 22)
        for k, ln in enumerate(lines):
            cv.drawString(x + 16 if k else x + 8, y, ("• " if k == 0 else "") + ln)
            y -= 8.2
        y -= 1.5
    cv.setStrokeColor(RED)
    cv.setLineWidth(0.6)
    cv.rect(x, y + 2, width, top - y - 2)


def export(study):
    info = build(study)
    return render(info), info


if __name__ == "__main__":
    if "--fields" in sys.argv:
        for (f, i), (p, rect) in sorted(_widgets(PdfReader(FORM_PATH)).items(), key=lambda kv: (kv[1][0], -kv[1][1][1])):
            print(p + 1, f, i, [round(v) for v in rect])

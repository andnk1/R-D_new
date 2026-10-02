"""
study_report.py - the R&D Study deliverable.

Builds the client-facing R&D study (cover, letter, credit summary, company and project
overview, QRE detail pages, gross receipts) from the screen answers, using the same
engine as the feasibility app (engine/rd_credit_calculator_v2.py) for every number.
Layout, CSS and wording come from rd_study_generator_Regular Credit Method.py.

Also keeps the preparer's uploaded documents (PDF, JPG/PNG, TXT) and appends them,
unchanged and in upload order, to the end of the PDF package.
"""
import base64
import io
import os
import re
import uuid
from datetime import datetime
from pathlib import Path

from engine import rd_credit_calculator_v2 as calc
import study as S

HERE = os.path.dirname(os.path.abspath(__file__))

# ═══════════════════════════════════════════════════════════════════════════════
# WORDING  (from rd_study_generator_Regular Credit Method.py – edit here for all clients)
# ═══════════════════════════════════════════════════════════════════════════════
TOC = [
    "Credit for Increasing Research Activities Summary",
    "Company Overview",
    "Project Overview",
    "Wage QRE",
    "Supply QRE",
    "Computer / Cloud QRE",
    "Contractor US",
    "Contractor Foreign",
    "Gross Receipts",
]

LETTER_PARAS = [
    "[study_date]",
    "EIN: [EIN_number]",
    "Re: [Company_Name] [tax_year] Research & Development Tax Credit Results",
    "Tax year [tax_year]",
    "Enclosed is the R&D tax credit calculation we have prepared for your review. This report provides detailed documentation of the research activities and related expenses undertaken by [Company_Name] for the [tax_year] in support of our claim for the Research and Development (R&D) Tax Credit under Section 41. This report complies with IRS requirements for filing a valid claim for refund.",
    "Table of Contents",
] + [f"{i}. {t}" for i, t in enumerate(TOC, 1)]

CREDIT_TAB_NARRATIVE = [
    "For the tax year [tax_year], [company_name], a [entity_type], completed a review of its qualified research activities and related expenditures for purposes of computing the Credit for Increasing Research Activities under IRC Section 41.",
    "The calculation was prepared based on qualified research expenditures such as wages for qualified services of [qre_wages], supplies of [qre_supplies], computer rental or lease costs of [qre_computers], U.S. contract research of [qre_contract_US], foreign contract research of [qre_contract_foreign], and basic research payments of [basic_research_payments]. Supporting details for the calculations are presented in the text sections. The calculation reflects the Section 280C reduced credit election, as shown on the table above.",
    "Based on the analysis performed, [company_name] computed an estimated federal R&D tax credit of [rd_credit_amount] for the [tax_year] tax year.",
    "This study was prepared on [study_date] by CFO Associates and is intended to summarize the qualified cost categories, methodology, and assumptions used in the credit calculation. The final credit should be supported by company books, payroll records, project documentation, and other relevant business records.",
]
# Used instead of paragraph 3 when the engine could not compute a credit
NO_CREDIT_PARA = ("Based on the analysis performed, a federal R&D tax credit could not be computed for the "
                  "[tax_year] tax year from the data provided. See the table above.")

DOLLAR_FIELDS = {"qre_wages", "qre_supplies", "qre_computers", "qre_contract_us", "qre_contract_foreign",
                 "basic_research_payments", "rd_credit_amount"}

# Uploaded documents
ALLOWED = {".pdf": "PDF", ".jpg": "Picture", ".jpeg": "Picture", ".png": "Picture", ".txt": "Text"}
ACCEPT_ATTR = ".pdf,.jpg,.jpeg,.png,.txt"


# ═══════════════════════════════════════════════════════════════════════════════
# HELPERS (same formatting as the generator)
# ═══════════════════════════════════════════════════════════════════════════════
def esc(text):
    return (str(text).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def fmt_dollar(v):
    n = v or 0.0
    return "$ -" if n == 0 else f"${n:,.2f}"


def fmt_dollar_total(v):
    return f"${(v or 0.0):,.2f}"


def fmt_pct(v):
    n = S.num(v)
    if n is None:
        return ""
    return f"{int(n)}%" if n == int(n) else f"{n:.1f}%"


def table_value(v):
    """Credit-table cell: '$ 1,234.56', '$ -' for zero (calc.fmt)."""
    return calc.fmt(v)


def apply_wildcards(text, fields):
    def replacer(m):
        key = m.group(1).lower()
        val = fields.get(key)
        if val is None:
            return m.group(0)
        if key in DOLLAR_FIELDS:
            try:
                return f"${float(val):,.2f}"
            except (TypeError, ValueError):
                return str(val)
        if isinstance(val, float):
            return str(int(val)) if val == int(val) else str(val)
        return str(val)
    return re.sub(r'\[([A-Za-z_]+)\]', replacer, text)


def text_paragraphs(text):
    """Each non-blank line of the screen text becomes one paragraph (as in the sample study)."""
    return [ln.strip() for ln in (text or "").splitlines() if ln.strip()]


# ═══════════════════════════════════════════════════════════════════════════════
# CSS  (unchanged from the generator)
# ═══════════════════════════════════════════════════════════════════════════════
CSS = """
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: Verdana, Geneva, Tahoma, sans-serif; font-size: 9pt; color: #1a1a1a;
       background: #e8e8e8; padding: 1.2rem 0 2rem 0; }
.page { width: 8.5in; min-height: 11in; margin: 0 auto 1.2rem auto; background: #fff;
        box-shadow: 0 3px 16px rgba(0,0,0,.18); display: flex; flex-direction: column; padding: 0; }
@media print {
    body { background: #fff; padding: 0; }
    .page { box-shadow: none; margin: 0; page-break-after: always; page-break-inside: avoid; }
    .page:last-child { page-break-after: avoid; }
}
.page-header { display: flex; align-items: center; justify-content: space-between; padding: 0.35rem 0.9in;
               background: #003366; color: #fff; font-size: 7.5pt; font-weight: bold; letter-spacing: 0.3px; flex-shrink: 0; }
.page-header span { flex: 1; }
.page-header span:nth-child(2) { text-align: center; }
.page-header span:last-child   { text-align: right; }
.page-content { flex: 1; padding: 0.55in 0.9in 0.4in 0.9in; overflow: hidden; }
.page-footer { flex-shrink: 0; padding: 0 0.9in 0.3in 0.9in; }
.footer-rule { border: none; border-top: 1px solid #b8c8d8; margin-bottom: 0.25rem; }
.footer-text { font-size: 7pt; color: #888; text-align: center; letter-spacing: 0.5px; }
.cover-content { flex: 1; padding: 0.6in 0.9in 0.4in 0.9in; display: flex; flex-direction: column; }
.cover-logo { height: 264px; margin-bottom: 0.45in; object-fit: contain; }
.cover-title-block { text-align: center; border-bottom: 3px solid #003366; padding-bottom: 1.2rem; margin-bottom: 1.2rem; }
.cover-title-block h1 { font-family: Verdana, sans-serif; font-size: 18pt; font-weight: bold; color: #003366; letter-spacing: 0.5px; line-height: 1.2; }
.cover-title-block h2 { font-size: 40pt; font-weight: bold; color: #003366; margin-top: 0.3rem; }
.cover-bottom-info { margin-top: auto; text-align: right; font-size: 9pt; color: #555; line-height: 1.8; }
.cover-letter { flex: 1; }
.letter-date { font-size: 9pt; margin: 0.4rem 0; }
.letter-re   { font-size: 9pt; font-weight: bold; margin: 0.5rem 0 0.7rem; }
.cover-letter p { font-size: 9pt; line-height: 1.65; margin-bottom: 0.6rem; text-align: justify; color: #1a1a1a; }
.section-title { font-family: Verdana, sans-serif; font-size: 13pt; font-weight: bold; color: #003366;
                 margin-bottom: 0.1rem; padding-bottom: 0.3rem; border-bottom: 2px solid #003366; }
.no-data { font-style: italic; color: #666; margin-top: 0.6rem; font-size: 8.5pt; }
.section-narrative { margin-top: 0.55rem; font-size: 8.5pt; line-height: 1.65; color: #1a1a1a; text-align: justify; }
.section-narrative p { margin-bottom: 0.45rem; }
table { width: 100%; border-collapse: collapse; font-family: Verdana, sans-serif; font-size: 8pt; margin-top: 0.5rem; }
thead tr { background-color: #003366; color: #fff; }
th { text-align: left; padding: 0.32rem 0.5rem; border: 1px solid #003366; font-weight: bold; font-size: 7.5pt; letter-spacing: 0.2px; }
td { padding: 0.27rem 0.5rem; border: 1px solid #ccd8e8; vertical-align: top; }
tr:nth-child(even) td { background-color: #f4f7fb; }
.total-row td { background-color: #dbe4f0 !important; border-top: 2px solid #003366; border-bottom: 2px solid #003366; font-weight: bold; }
.num { text-align: right; font-variant-numeric: tabular-nums; }
"""


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE BUILDERS
# ═══════════════════════════════════════════════════════════════════════════════
def _page_footer():
    return ('<div class="page-footer"><hr class="footer-rule">'
            '<div class="footer-text">CFO Associates</div></div>')


def _page_header(fields):
    return ('<div class="page-header">'
            '<span>' + esc(fields.get("company_name", "")) + '</span>'
            '<span>Research and Development Credit Study</span>'
            '<span>Tax Year ' + esc(str(fields.get("tax_year", ""))) + '</span></div>')


def section_page(title, body_html, fields):
    return ('<div class="page">' + _page_header(fields) + '<div class="page-content">'
            '<div class="section-title">' + esc(title) + '</div>' + body_html + '</div>'
            + _page_footer() + '</div>\n')


def no_data():
    return '<p class="no-data">No Data Provided.</p>'


def build_cover_pages(fields, logo_b64):
    company = esc(fields.get("company_name", ""))
    tax_year = esc(str(fields.get("tax_year", "")))
    preparer = esc(str(fields.get("preparer_name") or ""))
    logo = f'<img src="{logo_b64}" class="cover-logo" alt="CFO Associates">' if logo_b64 else ""
    page1 = ('<div class="page"><div class="cover-content">' + logo +
             '<div class="cover-title-block"><h1>Research and Development<br>Tax Credit</h1>'
             '<h2>' + company + '</h2></div>'
             '<div class="cover-bottom-info"><div>Tax Year Ended ' + esc(fields["year_end"]) + '</div>'
             + ('<div>Completed by ' + preparer + '</div>' if preparer else '') +
             '</div></div>' + _page_footer() + '</div>\n')

    html = '<div class="cover-content"><div class="cover-letter">'
    for idx, para in enumerate(LETTER_PARAS):
        if para.startswith("EIN:") and not fields.get("ein_number"):
            continue
        e = esc(apply_wildcards(para, fields))
        if para.startswith("Re:"):
            html += '<p class="letter-re">' + e + '</p>\n'
        elif idx == 0:
            html += '<p class="letter-date">' + e + '</p>\n'
        else:
            html += '<p>' + e + '</p>\n'
    html += '</div></div>'
    page2 = '<div class="page">\n' + html + _page_footer() + '</div>\n'
    return page1 + page2


def credit_rows(result, used, fields):
    """(label, value) rows of the 'Credit for Increasing Research Activities' table."""
    rec_label = "Recommended R&D Tax Credit"
    if "selected by preparer" in (result.get("recommended_note") or ""):
        rec_label = f"R&D Tax Credit — {S.METHOD_NAMES.get(used, '')} (selected by preparer)"
    return [
        ("Company Name", result["company_name"]),
        ("Tax Year", str(result["tax_year"])),
        ("Entity Type", result["entity_type"]),
        ("Are you electing the reduced credit under section 280C?", result["electing_280c"]),
        ("Wages for qualified services", table_value(result["qre_wages"])),
        ("Cost of supplies", table_value(result["qre_supplies"])),
        ("Rental or lease costs of computers", table_value(result["qre_computers"])),
        ("Applicable % of contract research expenses (65%)", table_value(result["qualified_contract_US"])),
        ("Total qualified research expenses", table_value(result["ordinary_qre"])),
        ("50% of total qualified research expenses", table_value(result["fifty_pct_qre"])),
        ("Regular Credit Method — 280C applied (15.8%)",
         calc.credit_cell(result["regular_credit_280c"], result["regular_method_status"])),
        ("Alternative Simplified Credit (ASC) — 280C applied",
         calc.credit_cell(result["asc_credit_280c"], result["asc_method_status"])),
        (rec_label, calc.credit_cell(result["recommended_credit"], result["recommended_note"])),
    ]


def narrative(fields):
    out = []
    for i, para in enumerate(CREDIT_TAB_NARRATIVE):
        if i == 2 and fields.get("rd_credit_amount") is None:
            para = NO_CREDIT_PARA
        out.append(apply_wildcards(para, fields))
    return out


def build_credit_tab_page(rows, paras, fields):
    body = '<table><tbody>'
    for label, value in rows:
        body += '<tr><td style="width:75%">' + esc(label) + '</td><td>' + esc(value) + '</td></tr>'
    body += '</tbody></table><div class="section-narrative">'
    body += "".join('<p>' + esc(p) + '</p>' for p in paras) + '</div>'
    return section_page("Credit for Increasing Research Activities", body, fields)


def build_narrative_page(title, paragraphs, fields):
    if not paragraphs:
        return section_page(title, no_data(), fields)
    html = '<div class="section-narrative">' + "".join(
        '<p>' + esc(apply_wildcards(p, fields)) + '</p>' for p in paragraphs) + '</div>'
    return section_page(title, html, fields)


def _lines(study, cat):
    """Rows entered on step 4 for one category, or one 'total only' row."""
    e = study["expenses"][cat]
    amt, mode = S.category_amount(study, cat)
    if mode == "total":
        return [{"name": "Total entered (no line detail)", "col2": "", "state": "",
                 "amount": S.num(e.get("total_amount")) or 0.0,
                 "pct": e.get("total_pct", "") if S.CATEGORIES[cat]["pct"] else "100",
                 "qual": amt}], amt
    out = []
    for r in e.get("rows", []):
        a = S.num(r.get("amount"))
        if a is None and not (r.get("name") or "").strip():
            continue
        p = (S.num(r.get("pct")) or 0.0) if S.CATEGORIES[cat]["pct"] else 100.0
        out.append({**r, "amount": a or 0.0, "qual": (a or 0.0) * p / 100.0})
    return out, amt


def build_wage_page(study, fields):
    rows, total = _lines(study, "wages")
    if not rows:
        return section_page("Wage QRE", no_data(), fields)
    html = ('<table><thead><tr><th>Employee Name</th><th>Job Title</th><th>State</th>'
            '<th class="num">Taxable Wages</th><th class="num">Qualified %</th><th class="num">QRE Amount</th>'
            '</tr></thead><tbody>')
    taxable = 0.0
    for r in rows:
        taxable += r["amount"]
        html += ('<tr><td>' + esc(r["name"]) + '</td><td>' + esc(r.get("col2", "")) + '</td><td>' + esc(r.get("state", "")) + '</td>'
                 '<td class="num">' + fmt_dollar(r["amount"]) + '</td><td class="num">' + esc(fmt_pct(r.get("pct"))) + '</td>'
                 '<td class="num">' + fmt_dollar(r["qual"]) + '</td></tr>')
    html += ('<tr class="total-row"><td colspan="3">Total</td><td class="num">' + fmt_dollar_total(taxable) + '</td><td></td>'
             '<td class="num">' + fmt_dollar_total(total) + '</td></tr></tbody></table>')
    return section_page("Wage QRE", html, fields)


def build_amount_page(title, study, cat, fields):
    """Supplies and Computer / Cloud – amount used for R&D (no qualified %)."""
    rows, total = _lines(study, cat)
    if not rows:
        return section_page(title, no_data(), fields)
    html = ('<table><thead><tr><th>Vendor</th><th>State</th><th>Type</th>'
            '<th class="num">Amount Used for R&amp;D</th><th class="num">Qualified Amount</th></tr></thead><tbody>')
    for r in rows:
        html += ('<tr><td>' + esc(r["name"]) + '</td><td>' + esc(r.get("state", "")) + '</td><td>' + esc(r.get("col2", "")) + '</td>'
                 '<td class="num">' + fmt_dollar(r["amount"]) + '</td><td class="num">' + fmt_dollar(r["qual"]) + '</td></tr>')
    html += ('<tr class="total-row"><td colspan="4">Total</td><td class="num">' + fmt_dollar_total(total)
             + '</td></tr></tbody></table>')
    return section_page(title, html, fields)


def build_contractor_page(title, study, cat, loc_label, fields):
    rows, total = _lines(study, cat)
    if not rows:
        return section_page(title, no_data(), fields)
    html = ('<table><thead><tr><th>Contractor Name</th><th>' + loc_label + '</th>'
            '<th class="num">Amount</th><th class="num">Qualified %</th>'
            '<th>Work Performed</th><th class="num">Total</th></tr></thead><tbody>')
    for r in rows:
        html += ('<tr><td>' + esc(r["name"]) + '</td><td>' + esc(r.get("state", "")) + '</td>'
                 '<td class="num">' + fmt_dollar(r["amount"]) + '</td><td class="num">' + esc(fmt_pct(r.get("pct"))) + '</td>'
                 '<td>' + esc(r.get("col2", "")) + '</td><td class="num">' + fmt_dollar(r["qual"]) + '</td></tr>')
    html += ('<tr class="total-row"><td colspan="5">Total</td><td class="num">' + fmt_dollar_total(total)
             + '</td></tr></tbody></table>')
    return section_page(title, html, fields)


def build_receipts_page(study, fields):
    years = [y for y in S.history_years(study) if S.num(study["history"]["gr"].get(str(y))) is not None]
    if not years:
        return section_page("Gross Receipts", no_data(), fields)
    html = '<table><thead><tr><th>Tax Year</th><th class="num">Gross Receipts</th></tr></thead><tbody>'
    for y in years:
        html += ('<tr><td>' + str(y) + '</td><td class="num">'
                 + fmt_dollar(S.num(study["history"]["gr"].get(str(y)))) + '</td></tr>')
    html += '</tbody></table>'
    return section_page("Gross Receipts", html, fields)


# ═══════════════════════════════════════════════════════════════════════════════
# RUN
# ═══════════════════════════════════════════════════════════════════════════════
def _fields(study, data, result, used):
    c = study["company"]
    ty = S.tax_year(study)
    fye = int(c.get("fye") or 12)
    end_year = ty if fye == 12 else ty + 1
    last_day = [31, 29 if end_year % 4 == 0 else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][fye - 1]
    month = datetime(2000, fye, 1).strftime("%B")
    return {
        "company_name": result["company_name"], "tax_year": ty, "entity_type": result["entity_type"],
        "ein_number": c.get("ein") or "", "preparer_name": c.get("preparer") or "",
        "study_date": data["study_date"], "year_end": f"{month} {last_day}, {end_year}",
        "qre_wages": result["qre_wages"], "qre_supplies": result["qre_supplies"],
        "qre_computers": result["qre_computers"], "qre_contract_us": data["qre_contract_US"],
        "qre_contract_foreign": data["qre_contract_foreign"], "basic_research_payments": 0.0,
        "rd_credit_amount": result["recommended_credit"], "method_used": used or "",
    }


def html_to_pdf(html_path, pdf_path):
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        return False, "Playwright is not installed."
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch()
            page = browser.new_page()
            page.goto(Path(html_path).resolve().as_uri())
            page.wait_for_load_state("networkidle")
            page.pdf(path=pdf_path, format="Letter", print_background=True,
                     margin={"top": "0", "bottom": "0", "left": "0", "right": "0"})
            browser.close()
        return True, ""
    except Exception as e:          # PDF is optional - the HTML study is still available
        return False, str(e).splitlines()[0][:200]


def run_rd_study(study, folder):
    data, result, used = S.calculate(study)
    calc.save_xlsx(result, data, os.path.join(folder, "rd_credit_calculator_v2_scenario_1.xlsx"))

    fields = _fields(study, data, result, used)
    rows = credit_rows(result, used, fields)
    paras = narrative(fields)
    p = study["projects"]

    logo_b64 = ""
    logo = os.path.join(HERE, "static", "logo.png")
    if os.path.exists(logo):
        with open(logo, "rb") as f:
            logo_b64 = "data:image/png;base64," + base64.b64encode(f.read()).decode()

    project_paras = text_paragraphs(p.get("description"))
    head = " — ".join(x for x in (p.get("name"), p.get("type")) if x)
    if head and project_paras:
        project_paras = [head] + project_paras

    pages = (build_cover_pages(fields, logo_b64)
             + build_credit_tab_page(rows, paras, fields)
             + build_narrative_page("Company Overview", text_paragraphs(p.get("company_description")), fields)
             + build_narrative_page("Project Overview", project_paras, fields)
             + build_wage_page(study, fields)
             + build_amount_page("Supply QRE", study, "supplies", fields)
             + build_amount_page("Computer / Cloud QRE", study, "cloud", fields)
             + build_contractor_page("Contractor US", study, "contract", "State", fields)
             + build_contractor_page("Contractor Foreign", study, "foreign", "Country", fields)
             + build_receipts_page(study, fields))

    company = result["company_name"]
    html = ("<!DOCTYPE html>\n<html lang=\"en\">\n<head>\n<meta charset=\"UTF-8\">\n"
            "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">\n"
            f"<title>R&amp;D Tax Credit Study — {esc(company)} {fields['tax_year']}</title>\n"
            "<style>\n" + CSS + "\n</style>\n</head>\n<body>\n" + pages + "</body>\n</html>")
    html_path = os.path.join(folder, "study.html")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)
    pdf_path = os.path.join(folder, "study.pdf")
    if os.path.exists(pdf_path):
        os.remove(pdf_path)
    pdf_ok, pdf_error = html_to_pdf(html_path, pdf_path)

    method, other = S.method_summary(result, used)
    return {
        "ran_at": datetime.now().strftime("%b %d, %Y %I:%M %p"),
        "company": company, "safe_name": re.sub(r"[^A-Za-z0-9_]", "_", company),
        "tax_year": result["tax_year"], "entity": result["entity_type"],
        "recommended": result["recommended_credit"], "method": method, "warnings": result["warnings"],
        "method_used": used, "other": other, "rows": rows, "narrative": paras,
        "pdf_ok": pdf_ok, "pdf_error": pdf_error,
    }


# ═══════════════════════════════════════════════════════════════════════════════
# UPLOADED DOCUMENTS + PDF PACKAGE
# ═══════════════════════════════════════════════════════════════════════════════
def _updir(folder):
    d = os.path.join(folder, "uploads")
    os.makedirs(d, exist_ok=True)
    return d


def add_uploads(study, folder, files):
    """Save uploaded files (no reading beyond a quick check that they open). -> (added, problems)"""
    docs = study.setdefault("uploads", [])
    added, problems = 0, []
    for fs in files:
        name = os.path.basename(fs.filename)
        ext = os.path.splitext(name)[1].lower()
        if ext not in ALLOWED:
            hint = " Save Word files as PDF first (File → Save As → PDF)." if ext in (".doc", ".docx") else ""
            problems.append(f"{name}: file type not accepted – use PDF, JPG, PNG or TXT.{hint}")
            continue
        raw = fs.read()
        if not raw:
            problems.append(f"{name}: the file is empty.")
            continue
        try:
            _to_pdf_pages(raw, ext)          # make sure it can be added to the package
        except Exception:
            problems.append(f"{name}: could not be opened (damaged or password-protected?).")
            continue
        doc_id = uuid.uuid4().hex[:12]
        stored = doc_id + ext
        with open(os.path.join(_updir(folder), stored), "wb") as f:
            f.write(raw)
        docs.append({"id": doc_id, "name": name, "file": stored, "type": ALLOWED[ext],
                     "size": f"{len(raw) / 1024:,.0f} KB" if len(raw) < 1024 * 1024 else f"{len(raw) / 1048576:,.1f} MB"})
        added += 1
    return added, problems


def change_upload(study, folder, doc_id, action):
    docs = study.get("uploads", [])
    i = next((n for n, d in enumerate(docs) if d["id"] == doc_id), None)
    if i is None:
        return
    if action == "remove":
        d = docs.pop(i)
        try:
            os.remove(os.path.join(_updir(folder), d["file"]))
        except OSError:
            pass
    elif action == "up" and i > 0:
        docs[i - 1], docs[i] = docs[i], docs[i - 1]
    elif action == "down" and i < len(docs) - 1:
        docs[i + 1], docs[i] = docs[i], docs[i + 1]


def _to_pdf_pages(raw, ext):
    """Return a PdfReader for one uploaded file (PDF as is; picture / text placed on Letter pages)."""
    from pypdf import PdfReader
    if ext == ".pdf":
        r = PdfReader(io.BytesIO(raw))
        if r.is_encrypted:
            r.decrypt("")
        len(r.pages)
        return r

    from reportlab.lib.pagesizes import letter
    from reportlab.pdfgen import canvas
    buf = io.BytesIO()
    W, H = letter
    if ext in (".jpg", ".jpeg", ".png"):
        from PIL import Image, ImageOps
        from reportlab.lib.utils import ImageReader
        img = ImageOps.exif_transpose(Image.open(io.BytesIO(raw)))
        if img.mode not in ("RGB", "L"):
            bg = Image.new("RGB", img.size, "white")
            bg.paste(img, mask=img.convert("RGBA").split()[-1])
            img = bg
        m = 36
        scale = min((W - 2 * m) / img.width, (H - 2 * m) / img.height, 1.0)   # fit the page, never enlarge
        w, h = img.width * scale, img.height * scale
        cv = canvas.Canvas(buf, pagesize=letter)
        cv.drawImage(ImageReader(img), (W - w) / 2, H - m - h, w, h)
        cv.showPage()
        cv.save()
    else:   # .txt
        from reportlab.lib.styles import ParagraphStyle
        from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer
        try:
            text = raw.decode("utf-8-sig")
        except UnicodeDecodeError:
            text = raw.decode("cp1252", errors="replace")
        st = ParagraphStyle("t", fontName="Helvetica", fontSize=10, leading=14)
        story = []
        for line in text.splitlines():
            story.append(Paragraph(esc(line).replace("\t", "&nbsp;" * 4), st) if line.strip() else Spacer(1, 8))
        doc = SimpleDocTemplate(buf, pagesize=letter, leftMargin=54, rightMargin=54, topMargin=54, bottomMargin=54)
        doc.build(story or [Spacer(1, 1)])
    return PdfReader(io.BytesIO(buf.getvalue()))


def build_package(study, folder):
    """Study PDF + uploaded documents (in the order shown on screen). -> (pdf bytes, skipped notes)"""
    from pypdf import PdfReader, PdfWriter
    pdf_path = os.path.join(folder, "study.pdf")
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(pdf_path)
    w = PdfWriter()
    w.append(PdfReader(pdf_path))
    skipped = []
    for d in study.get("uploads", []):
        try:
            with open(os.path.join(_updir(folder), d["file"]), "rb") as f:
                raw = f.read()
            w.append(_to_pdf_pages(raw, os.path.splitext(d["file"])[1].lower()))
        except Exception:
            skipped.append(f"{d['name']} could not be added to the package and was skipped.")
    out = io.BytesIO()
    w.write(out)
    return out.getvalue(), skipped

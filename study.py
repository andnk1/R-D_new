"""
study.py - turns what the preparer typed on the six screens into the inputs the existing
calculation engine expects, runs the engine, and builds the client-facing study.

The engine files in engine/ are copies of rd_credit_calculator_v2.py and
rd_feasibility_generator.py and are NOT modified. This file only feeds them data.
"""
import base64
import os
import re
from datetime import datetime
from pathlib import Path

from engine import rd_credit_calculator_v2 as calc
from engine import rd_feasibility_generator as gen

HERE = os.path.dirname(os.path.abspath(__file__))

STEPS = [
    ("company",  "Company & Filing Profile"),
    ("projects", "Projects"),
    ("history",  "Gross Receipts & Prior-Year QREs"),
    ("expenses", "Qualified R&D Expenses"),
    ("other",    "Other Considerations"),
    ("review",   "Review & Study"),
]

ENTITY_TYPES = {
    "C":  ("C Corporation", "C Corporation (Form 1120)"),
    "S":  ("S Corporation", "S Corporation (Form 1120-S)"),
    "P":  ("Partnership", "Partnership (Form 1065)"),
    "SP": ("Sole Proprietor / Single-Owner Pass-Through",
           "Sole Proprietor / Single-Owner Pass-Through (Schedule C)"),
}

# Expense categories: key -> settings for the screen and the workbook
CATEGORIES = {
    "wages": dict(tab="Wages", pct=True, sheet="Wages QRE",
                  cols=["Employee name", "Job title", "State", "Total taxable wages ($)", "Qualified %", "Qualified wages ($)"],
                  total_label="Total taxable wages ($)"),
    "supplies": dict(tab="Supplies", pct=False, sheet="Supplies QRE",
                     cols=["Vendor name", "Type", "State", "Amount used for R&D ($)", "Qualified supplies ($)"],
                     total_label="Total used for R&D ($)"),
    "cloud": dict(tab="Computer / Cloud Compute", pct=False, sheet="Cloud_Compute QRE",
                  cols=["Vendor name", "Type", "State", "Amount used for R&D ($)", "Qualified computer / cloud ($)"],
                  total_label="Total used for R&D ($)"),
    "contract": dict(tab="U.S. Contract Research", pct=True, sheet="Contractor US QRE",
                     cols=["Contractor name", "Work performed", "State", "Amount paid ($)", "Qualified %", "Qualified payment ($)"],
                     total_label="Total amount paid ($)"),
    "foreign": dict(tab="Foreign Contractors / Foreign R&E", pct=True, sheet="Contractor Foreign QRE",
                    cols=["Contractor name", "Work performed", "Country", "Amount paid ($)", "R&E %", "Foreign R&E ($)"],
                    total_label="Total amount paid ($)"),
}
MIN_ROWS = 5


# ---------------------------------------------------------------------------
# Blank study + number helpers
# ---------------------------------------------------------------------------

def blank_study():
    return {
        "company": {"entity": "C", "tax_year": "2025", "fye": "12", "cg": "no", "year_started": "",
                    "filed": "no", "ext": "no", "timely": "yes", "version": "v1"},
        "projects": {},
        "history": {"gr": {}, "qre": {}, "pte": "no", "pte_years": []},
        "expenses": {k: {"rows": [], "total_amount": "", "total_pct": ""} for k in CATEGORIES},
        "other": {"show_payroll": "yes"},
        "completed": {},
        "run": None,
    }


def num(v):
    """'1,234.50' / '$1,234' / '50%' -> float ; blank -> None (blank is NOT zero)."""
    if v is None:
        return None
    s = str(v).strip().replace("$", "").replace(",", "").replace("%", "").strip()
    if s == "":
        return None
    try:
        return float(s)
    except ValueError:
        return None


def money(v, cents=False):
    if v is None:
        return "—"
    return f"${v:,.2f}" if cents else f"${v:,.0f}"


def tax_year(study):
    try:
        return int(study["company"].get("tax_year") or 2025)
    except ValueError:
        return 2025


def year_started(study):
    n = num(study["company"].get("year_started"))
    return int(n) if n is not None else None


def history_years(study):
    """Study year back to the year operations began (max 10 prior years)."""
    ty, ys = tax_year(study), year_started(study)
    first = ty - 10
    if ys is not None and ty - 10 <= ys <= ty:
        first = ys
    return list(range(ty, first - 1, -1))


def first_gross_receipts_year(study):
    """Earliest year in the table with gross receipts above $0 (None = none yet)."""
    years = sorted(int(y) for y, v in study["history"]["gr"].items()
                   if (num(v) or 0) > 0 and int(y) in history_years(study))
    return years[0] if years else None


# ---------------------------------------------------------------------------
# Expense totals (same arithmetic as the screen)
# ---------------------------------------------------------------------------

def category_amount(study, cat):
    """Qualified amount used for a category (total-only line replaces the lines)."""
    e = study["expenses"][cat]
    has_pct = CATEGORIES[cat]["pct"]
    ta, tp = e.get("total_amount", ""), e.get("total_pct", "")
    if str(ta).strip() or (has_pct and str(tp).strip()):
        a = num(ta) or 0.0
        p = (num(tp) or 0.0) if has_pct else 100.0
        return a * p / 100.0, "total"
    total = 0.0
    for r in e.get("rows", []):
        a = num(r.get("amount"))
        if a is None:
            continue
        p = (num(r.get("pct")) or 0.0) if has_pct else 100.0
        total += a * p / 100.0
    return total, "lines"


def line_count(study, cat):
    return sum(1 for r in study["expenses"][cat].get("rows", [])
               if (r.get("name") or "").strip() or num(r.get("amount")) is not None)


# ---------------------------------------------------------------------------
# Build the engine input (same keys as the Calculation_Data sheet)
# ---------------------------------------------------------------------------

def engine_data(study):
    c, h = study["company"], study["history"]
    ty = tax_year(study)

    def year(v):
        n = num(v)
        return int(n) if n is not None else None

    data = {
        "company_name":          c.get("name") or "Company",
        "entity_type":           ENTITY_TYPES.get(c.get("entity"), ENTITY_TYPES["C"])[0],
        "tax_year":              ty,
        "EIN_number":            c.get("ein") or None,
        "study_date":            datetime.now().strftime("%B %d, %Y"),
        "preparer_name":         c.get("preparer") or None,
        "year_started":          year(c.get("year_started")),
        "year_research_started": year(c.get("year_research_started")),
        "company_industry":      c.get("industry") or None,
        "scenario_number":       1,
        "qre_wages":             category_amount(study, "wages")[0],
        "qre_supplies":          category_amount(study, "supplies")[0],
        "qre_computers":         category_amount(study, "cloud")[0],
        "qre_contract_US":       category_amount(study, "contract")[0],   # engine applies 65%
        "qre_contract_foreign":  category_amount(study, "foreign")[0],    # excluded by engine
        "annual_gross_receipts_current": num(h["gr"].get(str(ty))),
        "return_filed":          c.get("filed", "no"),
    }
    ys = year_started(study)
    for i in range(1, 11):
        y = ty - i
        if ys is not None and y < ys:
            # company did not exist yet -> confirmed zero (these rows are not shown on step 3)
            data[f"qre_yr_minus{i}"] = 0.0
            data[f"annual_gross_receipts_yr_minus{i}"] = 0.0
        else:
            data[f"qre_yr_minus{i}"] = num(h["qre"].get(str(y)))
            data[f"annual_gross_receipts_yr_minus{i}"] = num(h["gr"].get(str(y)))
    return data


# ---------------------------------------------------------------------------
# "Items to review" shown on the Review page before running
# ---------------------------------------------------------------------------

def review_flags(study):
    c, h = study["company"], study["history"]
    ty = tax_year(study)
    flags = []
    if not (c.get("name") or "").strip():
        flags.append(("Company name", "not entered on step 1."))
    if num(h["gr"].get(str(ty))) is None:
        flags.append(("Current-year gross receipts",
                      f"{ty} is blank, so the payroll-tax (QSB) option cannot be tested and will not appear in the study."))
    if not (c.get("year_research_started") or "").strip():
        flags.append(("Year R&D activity began", "not entered on step 1 – the Regular Credit cannot be calculated without it."))
    ys = year_started(study)
    missing = [str(ty - i) for i in (1, 2, 3)
               if not (ys is not None and ty - i < ys) and num(h["qre"].get(str(ty - i))) is None]
    if missing:
        flags.append(("Prior-year QREs", f"blank for {', '.join(missing)} – the ASC needs all three prior years (enter 0 if confirmed zero)."))
    if c.get("cg") in ("yes", "unsure"):
        flags.append(("Controlled group", "the company may be under common ownership or control. Group calculations may be required and are not yet included in the engine."))
    if h.get("pte") == "yes":
        flags.append(("Prior payroll election", "the 5-year election limit is not yet included in the engine."))
    return flags


# ---------------------------------------------------------------------------
# Run the engine + build the client-facing study (HTML, PDF, workbook)
# ---------------------------------------------------------------------------

def run_study(study, folder):
    data = engine_data(study)
    result = calc.calculate(data)

    xlsx_path = os.path.join(folder, "rd_credit_calculator_v2_scenario_1.xlsx")
    calc.save_xlsx(result, data, xlsx_path)

    # Feasibility generator – same calls as its main(), pointed at this study's folder
    logo = os.path.join(HERE, "static", "logo.png")
    if os.path.exists(logo):
        with open(logo, "rb") as f:
            gen.LOGO_B64 = "data:image/png;base64," + base64.b64encode(f.read()).decode()
    fields = gen.load_company_info([xlsx_path])
    body = gen.build_cover_pages(fields) + gen.build_content_pages(fields)
    body += gen.build_scenario_page(gen.load_scenario(xlsx_path), fields, gen.load_gross_receipts(xlsx_path))
    company = fields.get("company_name", "Company")
    html = gen.wrap_html(body, title=f"R&D Tax Credit Feasibility Study - {company} {fields.get('tax_year', '')}")
    html_path = os.path.join(folder, "study.html")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)

    pdf_ok, pdf_error = False, ""
    if gen.HAS_PLAYWRIGHT:
        try:
            with gen.sync_playwright() as p:
                browser = p.chromium.launch()
                page = browser.new_page()
                page.goto(Path(html_path).resolve().as_uri())
                page.wait_for_load_state("networkidle")
                page.pdf(path=os.path.join(folder, "study.pdf"), format="Letter",
                         print_background=True, margin={"top": "0", "bottom": "0", "left": "0", "right": "0"})
                browser.close()
            pdf_ok = True
        except Exception as e:          # PDF is optional - the HTML study is still available
            pdf_error = str(e).splitlines()[0][:200]
    else:
        pdf_error = "Playwright is not installed."

    safe = re.sub(r"[^A-Za-z0-9_]", "_", company)
    rec = result["recommended_credit"]
    note = result.get("recommended_note", "")
    if rec is None:
        method = "Insufficient data — see notes below"
    elif note.startswith("ASC"):
        method = "Alternative Simplified Credit (ASC) — 280C applied (feasibility estimate)"
    else:
        method = "Regular Credit Method — 280C reduced credit (15.8%)"

    return {
        "ran_at": datetime.now().strftime("%b %d, %Y %I:%M %p"),
        "company": company, "safe_name": safe,
        "tax_year": result["tax_year"], "entity": result["entity_type"],
        "wages": result["qre_wages"], "supplies": result["qre_supplies"],
        "computers": result["qre_computers"], "contract65": result["qualified_contract_US"],
        "total": result["ordinary_qre"],
        "regular": result["regular_credit_280c"], "regular_status": result["regular_method_status"],
        "asc": result["asc_credit_280c"], "asc_status": result["asc_method_status"],
        "recommended": rec, "method": method, "warnings": result["warnings"],
        "pdf_ok": pdf_ok, "pdf_error": pdf_error,
    }

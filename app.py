"""
R&D Tax Credit – Feasibility Study web app (Flask)

Run locally:   python app.py        then open http://localhost:5000
On Railway:    gunicorn app:app     (see Dockerfile / README.md)
"""
import io
import os
from datetime import date

from flask import (Flask, abort, flash, redirect, render_template, request,
                   send_file, session, url_for)

import study as S
import workbook as W
from store import FileStore

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-only-change-me")
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024        # 20 MB uploads
store = FileStore()

STEP_URLS = {"company": "company", "projects": "projects", "history": "history",
             "expenses": "expenses", "other": "other", "review": "review"}


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

def current():
    """Return (study_id, study) for this browser; create one if needed."""
    sid = session.get("study_id")
    data = store.get(sid) if sid else {}
    if not data:
        sid = store.new_id()
        session["study_id"] = sid
        data = S.blank_study()
        store.save(sid, data)
    return sid, data


def finish(sid, data, step, next_step):
    data["completed"][step] = date.today().isoformat()
    data["run"] = None                       # inputs changed -> results must be re-run
    store.save(sid, data)
    if request.form.get("go") == "stay":
        flash("Saved.")
        return redirect(url_for(step))
    return redirect(url_for(next_step))


@app.context_processor
def inject():
    sid = session.get("study_id")
    data = store.get(sid) if sid else {}
    return {"steps": S.STEPS, "done": (data or {}).get("completed", {}),
            "meta_year": (data or {}).get("company", {}).get("tax_year", ""),
            "meta_version": (data or {}).get("company", {}).get("version", ""),
            "money": S.money}


@app.before_request
def _prune():
    if request.endpoint == "landing":
        store.prune()


# ---------------------------------------------------------------------------
# landing
# ---------------------------------------------------------------------------

@app.route("/")
def landing():
    sid = session.get("study_id")
    has_study = bool(sid and store.get(sid).get("completed"))
    return render_template("landing.html", has_study=has_study)


@app.route("/feasibility/new")
def new_study():
    old = session.pop("study_id", None)
    if old:
        store.delete(old)
    current()
    return redirect(url_for("company"))


# ---------------------------------------------------------------------------
# step 1 – company & filing profile
# ---------------------------------------------------------------------------

@app.route("/feasibility/company", methods=["GET", "POST"])
def company():
    sid, data = current()
    if request.method == "POST":
        f = request.form
        keys = ["name", "ein", "entity", "industry", "tax_year", "fye", "year_started",
                "year_research_started", "first_receipts_year", "cg", "filed", "ext",
                "timely", "preparer", "version"]
        data["company"] = {k: f.get(k, "").strip() for k in keys}
        return finish(sid, data, "company", "projects")
    return render_template("company.html", step="company", c=data["company"],
                           entities=S.ENTITY_TYPES, today=date.today().strftime("%b %d, %Y"))


# ---------------------------------------------------------------------------
# step 2 – project
# ---------------------------------------------------------------------------

@app.route("/feasibility/projects", methods=["GET", "POST"])
def projects():
    sid, data = current()
    if request.method == "POST":
        f = request.form
        p = {"company_description": f.get("company_description", "").strip(),
             "name": f.get("name", "").strip(), "type": f.get("type", "").strip(),
             "description": f.get("description", "").strip()}
        for field, upload in (("company_description", "company_file"), ("description", "project_file")):
            fs = request.files.get(upload)
            if fs and fs.filename:
                text = W.read_text(fs)
                if text:
                    p[field] = text
                else:
                    flash(f"Could not read {fs.filename}. Use a .docx or .txt file.", "err")
        data["projects"] = p
        return finish(sid, data, "projects", "history")
    return render_template("projects.html", step="projects", p=data["projects"])


# ---------------------------------------------------------------------------
# step 3 – gross receipts & prior-year QREs
# ---------------------------------------------------------------------------

@app.route("/feasibility/history", methods=["GET", "POST"])
def history():
    sid, data = current()
    years = S.history_years(data)
    if request.method == "POST":
        f = request.form
        h = data["history"]
        for y in years:
            h["gr"][str(y)] = f.get(f"gr_{y}", "").strip()
            if y != years[0]:
                h["qre"][str(y)] = f.get(f"qre_{y}", "").strip()
        h["prior"] = f.get("prior", "no")
        h["pte"] = f.get("pte", "no")
        h["pte_years"] = f.getlist("pte_years") if h["pte"] == "yes" else []
        return finish(sid, data, "history", "expenses")
    return render_template("history.html", step="history", h=data["history"], years=years)


# ---------------------------------------------------------------------------
# step 4 – qualified R&D expenses
# ---------------------------------------------------------------------------

def _read_expense_form(data):
    f = request.form
    for cat in S.CATEGORIES:
        n = int(f.get(f"{cat}-count", S.MIN_ROWS) or S.MIN_ROWS)
        rows = []
        for i in range(n):
            r = {k: f.get(f"{cat}-{i}-{k}", "").strip() for k in ("name", "col2", "state", "amount", "pct")}
            if any(r.values()):
                rows.append(r)
        data["expenses"][cat] = {"rows": rows,
                                 "total_amount": f.get(f"{cat}-total_amount", "").strip(),
                                 "total_pct": f.get(f"{cat}-total_pct", "").strip()}


@app.route("/feasibility/expenses", methods=["GET", "POST"])
def expenses():
    sid, data = current()
    if request.method == "POST":
        _read_expense_form(data)
        fs = request.files.get("workbook")
        if fs and fs.filename:
            try:
                found = W.parse_upload(fs)
            except Exception:
                found = None
            if not found:
                store.save(sid, data)
                flash("Could not read that workbook. Use the downloaded template (tabs and headers unchanged).", "err")
                return redirect(url_for("expenses", tab=request.form.get("tab", "wages")))
            else:
                for cat, rows in found.items():
                    data["expenses"][cat] = {"rows": rows, "total_amount": "", "total_pct": ""}
                data["run"] = None
                store.save(sid, data)
                read = ", ".join(f"{S.CATEGORIES[c]['tab']}: {len(r)} line(s)" for c, r in found.items())
                flash(f"Workbook read – {read}. Review the lines below, then save.")
                return redirect(url_for("expenses", tab=request.form.get("tab", "wages")))
        return finish(sid, data, "expenses", "other")
    rows = {}
    for cat in S.CATEGORIES:
        r = list(data["expenses"][cat]["rows"])
        while len(r) < S.MIN_ROWS:
            r.append({"name": "", "col2": "", "state": "", "amount": "", "pct": ""})
        rows[cat] = r
    return render_template("expenses.html", step="expenses", cats=S.CATEGORIES, rows=rows,
                           e=data["expenses"], tab=request.args.get("tab", "wages"))


@app.route("/feasibility/workbook")
def workbook_template():
    return send_file(io.BytesIO(W.template_bytes()), as_attachment=True,
                     download_name="RD_Feasibility_Expenses.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ---------------------------------------------------------------------------
# step 5 – other considerations
# ---------------------------------------------------------------------------

@app.route("/feasibility/other", methods=["GET", "POST"])
def other():
    sid, data = current()
    if request.method == "POST":
        data["other"] = {"show_payroll": request.form.get("show_payroll", "yes")}
        return finish(sid, data, "other", "review")
    return render_template("other.html", step="other", o=data["other"])


# ---------------------------------------------------------------------------
# step 6 – review & study
# ---------------------------------------------------------------------------

def _summary(data):
    c, p, h = data["company"], data["projects"], data["history"]
    ty = S.tax_year(data)
    months = ["January", "February", "March", "April", "May", "June", "July", "August",
              "September", "October", "November", "December"]
    fye = months[int(c.get("fye") or 12) - 1]
    filed = "Return filed" if c.get("filed") == "yes" else "Return not filed"
    ext = "extension filed" if c.get("ext") == "yes" else "no extension"
    rd = f"R&D began {c['year_research_started']}" if c.get("year_research_started") else "R&D start year not entered"
    s1 = f"{S.ENTITY_TYPES.get(c.get('entity'), S.ENTITY_TYPES['C'])[0]} · Tax year {ty} · Year-end {fye} · {rd} · {filed}, {ext}"
    s2 = " · ".join(x for x in (p.get("name"), p.get("type")) if x) or "No project entered"
    gr_years = sorted(int(y) for y, v in h["gr"].items() if S.num(v) is not None)
    qre_years = sorted(int(y) for y, v in h["qre"].items() if S.num(v) is not None)
    def span(ys):
        return f"{ys[0]}–{ys[-1]}" if len(ys) > 1 else (str(ys[0]) if ys else "none")
    s3 = f"Gross receipts entered: {span(gr_years)} · Prior-year QREs entered: {span(qre_years)}"
    if S.num(h["gr"].get(str(ty))) is None:
        s3 += f" · {ty} gross receipts blank"
    parts = []
    for cat, cfg in S.CATEGORIES.items():
        amt, mode = S.category_amount(data, cat)
        label = cfg["tab"].split(" / ")[0] if cat != "foreign" else "Foreign R&E"
        detail = "total only" if mode == "total" else f"{S.line_count(data, cat)} line(s)"
        txt = f"{label}: {detail} – {S.money(amt)}"
        if cat == "contract":
            txt += f" × 65% = {S.money(amt * 0.65)}"
        if cat == "foreign":
            txt += " (not in the credit)"
        parts.append(txt)
    s4 = " · ".join(parts)
    s5 = f"Show payroll option if eligible: {'Yes' if data['other'].get('show_payroll') == 'yes' else 'No'} · Section 280C reduced credit · Method determined by the app"
    return [("Company & Filing Profile", s1, "company"), ("Projects", s2, "projects"),
            ("Gross Receipts & Prior-Year QREs", s3, "history"), ("Qualified R&D Expenses", s4, "expenses"),
            ("Other Considerations", s5, "other")]


@app.route("/feasibility/review", methods=["GET", "POST"])
def review():
    sid, data = current()
    if request.method == "POST":
        data["run"] = S.run_study(data, store.folder(sid))
        data["completed"]["review"] = date.today().isoformat()
        store.save(sid, data)
        return redirect(url_for("review") + "#results")
    return render_template("review.html", step="review", summary=_summary(data),
                           flags=S.review_flags(data), run=data.get("run"))


FILES = {"html": ("study.html", "text/html"), "pdf": ("study.pdf", "application/pdf"),
         "xlsx": ("rd_credit_calculator_v2_scenario_1.xlsx",
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}


@app.route("/feasibility/file/<kind>")
def study_file(kind):
    sid, data = current()
    run = data.get("run")
    if kind not in FILES or not run:
        abort(404)
    name, mime = FILES[kind]
    path = os.path.join(store.folder(sid), name)
    if not os.path.exists(path):
        abort(404)
    ty = run.get("tax_year", "")
    nice = {"html": f"rd_feasibility_{run['safe_name']}_{ty}.html",
            "pdf": f"rd_feasibility_{run['safe_name']}_{ty}.pdf",
            "xlsx": f"R&D_Credit_{run['safe_name']}_1.xlsx"}[kind]
    return send_file(path, mimetype=mime, as_attachment=(kind != "html"), download_name=nice)


if __name__ == "__main__":
    app.run(debug=True, port=int(os.environ.get("PORT", 5000)))

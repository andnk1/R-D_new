"""
R&D Tax Credit web app (Flask) – two apps on the same six input screens:

  /feasibility/...   Feasibility Study  (existing feasibility generator, unchanged output)
  /study/...         R&D Study          (full study PDF + uploaded documents package)

Both use the same engine (engine/rd_credit_calculator_v2.py) and the same steps 1-5.
Only the final page (step 6) and the deliverable differ.

Run locally:   python app.py        then open http://localhost:5000
On Railway:    gunicorn app:app     (see Dockerfile / README.md)
"""
import io
import os
from datetime import date

from flask import (Blueprint, Flask, abort, flash, redirect, render_template, request,
                   send_file, session, url_for)

import form6765 as F
import study as S
import study_report as R
import workbook as W
from store import FileStore

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-only-change-me")
app.config["MAX_CONTENT_LENGTH"] = 60 * 1024 * 1024        # 60 MB per upload request
store = FileStore()

KINDS = {"feas": ("feasibility", "Feasibility Study"), "study": ("study", "R&D Study")}


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

def _kind():
    return request.blueprint or "feas"


def _sid_key(kind=None):
    # "study_id" kept for the feasibility app so studies already in progress keep working
    return "study_id" if (kind or _kind()) == "feas" else "study_id_rd"


def current():
    """Return (study_id, study) for this browser and app; create one if needed."""
    key = _sid_key()
    sid = session.get(key)
    data = store.get(sid) if sid else {}
    if not data:
        sid = store.new_id()
        session[key] = sid
        data = S.blank_study()
        data["kind"] = _kind()
        store.save(sid, data)
    return sid, data


def finish(sid, data, step, next_step):
    data["completed"][step] = date.today().isoformat()
    data["run"] = None                       # inputs changed -> results must be re-run
    data["method"] = ""                      # ... and the method goes back to the engine's pick
    store.save(sid, data)
    if request.form.get("go") == "stay":
        flash("Saved.")
        return redirect(url_for("." + step))
    return redirect(url_for("." + next_step))


@app.context_processor
def inject():
    kind = request.blueprint
    sid = session.get(_sid_key(kind)) if kind else None
    data = store.get(sid) if sid else {}
    return {"steps": S.STEPS, "done": (data or {}).get("completed", {}),
            "meta_year": (data or {}).get("company", {}).get("tax_year", ""),
            "meta_version": (data or {}).get("company", {}).get("version", ""),
            "money": S.money, "is_study": kind == "study",
            "app_name": KINDS.get(kind or "", ("", "R&D Tax Credit"))[1]}


@app.before_request
def _prune():
    if request.endpoint == "landing":
        store.prune()


@app.route("/")
def landing():
    has = {}
    for kind in KINDS:
        sid = session.get(_sid_key(kind))
        has[kind] = bool(sid and store.get(sid).get("completed"))
    return render_template("landing.html", has=has)


# ---------------------------------------------------------------------------
# The six screens – registered once for each app (blueprint)
# ---------------------------------------------------------------------------

def make_blueprint(kind):
    prefix, label = KINDS[kind]
    bp = Blueprint(kind, __name__, url_prefix="/" + prefix)

    @bp.route("/new")
    def new_study():
        old = session.pop(_sid_key(), None)
        if old:
            store.delete(old)
        current()
        return redirect(url_for(".company"))

    # -- step 1 – company & filing profile ---------------------------------
    @bp.route("/company", methods=["GET", "POST"])
    def company():
        sid, data = current()
        if request.method == "POST":
            f = request.form
            keys = ["name", "ein", "entity", "industry", "tax_year", "fye", "year_started",
                    "year_research_started", "cg", "filed", "ext",
                    "timely", "preparer", "version"]
            data["company"] = {k: f.get(k, "").strip() for k in keys}
            return finish(sid, data, "company", "projects")
        return render_template("company.html", step="company", c=data["company"],
                               entities=S.ENTITY_TYPES, today=date.today().strftime("%b %d, %Y"))

    # -- step 2 – project ------------------------------------------------------
    @bp.route("/projects", methods=["GET", "POST"])
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

    # -- step 3 – gross receipts & prior-year QREs -----------------------------
    @bp.route("/history", methods=["GET", "POST"])
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
            h["pte"] = f.get("pte", "no")
            h["pte_years"] = f.getlist("pte_years") if h["pte"] == "yes" else []
            return finish(sid, data, "history", "expenses")
        ys = S.year_started(data)
        if ys is None:
            note = "Enter “Year operations began” on step 1 to show only the years the company existed."
        elif ys > S.tax_year(data):
            note = f"“Year operations began” ({ys}) is after the study year – check step 1."
        elif S.tax_year(data) - 10 < ys:
            note = f"Showing {years[-1]}–{years[0]}. Years before {ys} are treated as zero – the company did not exist."
        else:
            note = ""
        return render_template("history.html", step="history", h=data["history"], years=years,
                               years_note=note, first_gr=S.first_gross_receipts_year(data))

    # -- step 4 – qualified R&D expenses ---------------------------------------
    @bp.route("/expenses", methods=["GET", "POST"])
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
                    return redirect(url_for(".expenses", tab=request.form.get("tab", "wages")))
                for cat, rows in found.items():
                    data["expenses"][cat] = {"rows": rows, "total_amount": "", "total_pct": ""}
                data["run"] = None
                store.save(sid, data)
                read = ", ".join(f"{S.CATEGORIES[c]['tab']}: {len(r)} line(s)" for c, r in found.items())
                flash(f"Workbook read – {read}. Review the lines below, then save.")
                return redirect(url_for(".expenses", tab=request.form.get("tab", "wages")))
            return finish(sid, data, "expenses", "other")
        rows = {}
        for cat in S.CATEGORIES:
            r = list(data["expenses"][cat]["rows"])
            while len(r) < S.MIN_ROWS:
                r.append({"name": "", "col2": "", "state": "", "amount": "", "pct": "", "officer": "No"})
            rows[cat] = r
        return render_template("expenses.html", step="expenses", cats=S.CATEGORIES, rows=rows,
                               e=data["expenses"], tab=request.args.get("tab", "wages"))

    @bp.route("/workbook")
    def workbook_template():
        name = "RD_Study_Expenses.xlsx" if kind == "study" else "RD_Feasibility_Expenses.xlsx"
        return send_file(io.BytesIO(W.template_bytes(label)), as_attachment=True, download_name=name,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # -- step 5 – other considerations -----------------------------------------
    @bp.route("/other", methods=["GET", "POST"])
    def other():
        sid, data = current()
        if request.method == "POST":
            data["other"] = {"show_payroll": request.form.get("show_payroll", "yes")}
            return finish(sid, data, "other", "review")
        return render_template("other.html", step="other", o=data["other"])

    # -- step 6 – review & study -------------------------------------------------
    def _run(data, sid):
        if kind == "study":
            return R.run_rd_study(data, store.folder(sid))
        return S.run_study(data, store.folder(sid))

    @bp.route("/review", methods=["GET", "POST"])
    def review():
        sid, data = current()
        if request.method == "POST":
            data["run"] = _run(data, sid)
            data["completed"]["review"] = date.today().isoformat()
            store.save(sid, data)
            return redirect(url_for(".review") + "#results")
        tpl = "review_study.html" if kind == "study" else "review.html"
        return render_template(tpl, step="review", summary=_summary(data),
                               flags=S.review_flags(data), run=data.get("run"),
                               uploads=data.get("uploads", []), accept=R.ACCEPT_ATTR)

    @bp.route("/method", methods=["POST"])
    def switch_method():
        """Re-run using the method chosen on the Review page (engine math unchanged)."""
        sid, data = current()
        m = request.form.get("method", "")
        data["method"] = m if m in ("regular", "asc") else ""
        data["run"] = _run(data, sid)
        store.save(sid, data)
        flash(f"Study prepared using the {S.METHOD_NAMES.get(data['run'].get('method_used'), 'recommended')} method.")
        return redirect(url_for(".review") + "#results")

    @bp.route("/form6765")
    def export_6765():
        """Draft Form 6765 – separate PDF, not part of the study."""
        sid, data = current()
        run = data.get("run")
        if not run or run.get("recommended") is None:
            flash("Run the study first – Form 6765 needs a calculated credit.", "err")
            return redirect(url_for(".review"))
        try:
            pdf, info = F.export(data)
        except Exception as e:
            flash(f"Form 6765 could not be created: {e}", "err")
            return redirect(url_for(".review") + "#results")
        name = f"Form_6765_DRAFT_{run['safe_name']}_{run.get('tax_year', '')}.pdf"
        return send_file(io.BytesIO(pdf), mimetype="application/pdf", as_attachment=True, download_name=name)

    @bp.route("/file/<what>")
    def study_file(what):
        sid, data = current()
        run = data.get("run")
        files = {"html": ("study.html", "text/html"), "pdf": ("study.pdf", "application/pdf"),
                 "xlsx": ("rd_credit_calculator_v2_scenario_1.xlsx",
                          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
        if what not in files or not run:
            abort(404)
        name, mime = files[what]
        path = os.path.join(store.folder(sid), name)
        if not os.path.exists(path):
            abort(404)
        ty = run.get("tax_year", "")
        base = "rd_study" if kind == "study" else "rd_feasibility"
        nice = {"html": f"{base}_{run['safe_name']}_{ty}.html",
                "pdf": f"{base}_{run['safe_name']}_{ty}.pdf",
                "xlsx": f"R&D_Credit_{run['safe_name']}_1.xlsx"}[what]
        return send_file(path, mimetype=mime, as_attachment=(what != "html"), download_name=nice)

    # -- R&D Study only: uploaded documents + PDF package ------------------------
    if kind == "study":
        @bp.route("/documents", methods=["POST"])
        def upload_documents():
            sid, data = current()
            files = [fs for fs in request.files.getlist("documents") if fs and fs.filename]
            if not files:
                flash("Choose one or more files to upload.", "err")
                return redirect(url_for(".review") + "#documents")
            added, problems = R.add_uploads(data, store.folder(sid), files)
            store.save(sid, data)
            if added:
                flash(f"{added} document(s) added to the package.")
            for p in problems:
                flash(p, "err")
            return redirect(url_for(".review") + "#documents")

        @bp.route("/documents/<doc_id>/<action>", methods=["POST"])
        def change_document(doc_id, action):
            sid, data = current()
            R.change_upload(data, store.folder(sid), doc_id, action)
            store.save(sid, data)
            return redirect(url_for(".review") + "#documents")

        @bp.route("/package")
        def package():
            sid, data = current()
            run = data.get("run")
            if not run:
                flash("Run the study first.", "err")
                return redirect(url_for(".review"))
            try:
                pdf, skipped = R.build_package(data, store.folder(sid))
            except FileNotFoundError:
                flash("The study PDF was not created, so the package cannot be built. See the note under the buttons.", "err")
                return redirect(url_for(".review") + "#results")
            for s in skipped:
                flash(s, "err")
            name = f"rd_study_package_{run['safe_name']}_{run.get('tax_year', '')}.pdf"
            return send_file(io.BytesIO(pdf), mimetype="application/pdf", as_attachment=True, download_name=name)

    return bp


def _read_expense_form(data):
    f = request.form
    for cat in S.CATEGORIES:
        n = int(f.get(f"{cat}-count", S.MIN_ROWS) or S.MIN_ROWS)
        rows = []
        for i in range(n):
            r = {k: f.get(f"{cat}-{i}-{k}", "").strip() for k in ("name", "col2", "state", "amount", "pct")}
            if any(r.values()) and S.CATEGORIES[cat].get("officer"):
                r["officer"] = f.get(f"{cat}-{i}-officer", "No")
            if any(v for k, v in r.items() if k != "officer"):
                rows.append(r)
        data["expenses"][cat] = {"rows": rows,
                                 "total_amount": f.get(f"{cat}-total_amount", "").strip(),
                                 "total_pct": f.get(f"{cat}-total_pct", "").strip()}


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
    fg = S.first_gross_receipts_year(data)
    s3 = (f"Gross receipts entered: {span(gr_years)} · Prior-year QREs entered: {span(qre_years)}"
          f" · First year with gross receipts: {fg or 'none yet'}")
    if S.num(h["gr"].get(str(ty))) is None:
        s3 += f" · {ty} gross receipts blank"
    parts = []
    for cat, cfg in S.CATEGORIES.items():
        amt, mode = S.category_amount(data, cat)
        lab = cfg["tab"].split(" / ")[0] if cat != "foreign" else "Foreign R&E"
        detail = "total only" if mode == "total" else f"{S.line_count(data, cat)} line(s)"
        txt = f"{lab}: {detail} – {S.money(amt)}"
        if cat == "wages":
            ow = S.officer_wages(data)
            txt += f" (officers {S.money(ow)})" if ow else ""
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


for _k in KINDS:
    app.register_blueprint(make_blueprint(_k))


if __name__ == "__main__":
    app.run(debug=True, port=int(os.environ.get("PORT", 5000)))

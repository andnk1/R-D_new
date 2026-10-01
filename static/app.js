// R&D Feasibility app – on-screen helpers only (live totals, show/hide, status lines).
// Nothing here is saved or calculated for the study; the server does that when you save / run.
(function () {
  const page = document.body.dataset.page;
  const $ = (s, el) => (el || document).querySelector(s);
  const $$ = (s, el) => Array.from((el || document).querySelectorAll(s));
  const money = n => "$" + Math.round(n).toLocaleString("en-US");
  const num = v => { const x = parseFloat(String(v).replace(/[,$\s%]/g, "")); return isNaN(x) ? 0 : x; };
  const val = name => { const e = $('input[name="' + name + '"]:checked'); return e ? e.value : null; };

  // ---------------- step 1: filing status line ----------------
  if (page === "company") {
    const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
    const fmt = d => MONTHS[d.getMonth()] + " " + d.getDate() + ", " + d.getFullYear();
    const roll = d => { while (d.getDay() === 0 || d.getDay() === 6) d.setDate(d.getDate() + 1); return d; };
    const m15 = (y, m, n) => roll(new Date(y, m - 1 + n, 15));
    function dueDates(entity, ty, fye) {
      const endY = (fye === 12) ? ty : ty + 1;
      if (entity === "C") {
        const june = (fye === 6 && ty < 2026);
        return { orig: m15(endY, fye, june ? 3 : 4), ext: m15(endY, fye, june ? 10 : 10) };
      }
      if (entity === "S" || entity === "P") return { orig: m15(endY, fye, 3), ext: m15(endY, fye, 9) };
      return { orig: m15(endY, fye, 4), ext: m15(endY, fye, 10) };
    }
    function refresh() {
      const entity = $("#entity").value, ty = parseInt($("#taxyear").value, 10), fye = parseInt($("#fye").value, 10);
      const { orig, ext } = dueDates(entity, ty, fye);
      const filed = val("filed"), extn = val("ext"), timely = val("timely");
      const today = new Date(); today.setHours(0, 0, 0, 0);
      $("#timelyq").style.display = (filed === "yes") ? "" : "none";
      $("#bbawarn").style.display = (entity === "P" && filed === "yes") ? "" : "none";
      $("#cgwarn").style.display = (val("cg") === "no") ? "none" : "";
      let msg, bad = false;
      if (filed !== "yes") {
        const deadline = (extn === "yes") ? ext : orig;
        const days = Math.round((deadline - today) / 86400000);
        if (today <= deadline) {
          msg = "<b>Timely original return still possible.</b> The " + (extn === "yes" ? "extended" : "original") +
            " due date is " + fmt(deadline) + " (" + days + " day" + (days === 1 ? "" : "s") + " from today). " +
            "Elections that require a timely original return – Section 280C and, if eligible, the payroll tax election – can still be preserved.";
          if (extn !== "yes" && today <= ext) msg += " Consider filing an extension to keep these options open until " + fmt(ext) + ".";
        } else {
          bad = true;
          msg = "<b>The timely-filing window has passed</b> (" + fmt(deadline) + "). The return can still be filed and an otherwise valid research credit may still be claimed, but elections that require a timely original return may be unavailable.";
        }
      } else if (timely === "yes") {
        msg = "<b>Return filed on time.</b> Any 280C or payroll election made on that return stands. A credit may still be claimed or increased on an amended return if the refund statute is open – but 280C and the payroll election cannot be first made by amendment.";
      } else {
        bad = true;
        msg = "<b>Return filed late.</b> The credit may still be claimed on an amended return if the refund statute is open, but elections that require a timely original return are generally unavailable.";
      }
      const box = $("#statusmsg");
      box.innerHTML = msg;
      box.style.background = bad ? "#fdf0f0" : ""; box.style.borderLeftColor = bad ? "#8b0000" : ""; box.style.color = bad ? "#8b0000" : "";
    }
    $$("select, input[type=radio]").forEach(e => e.addEventListener("change", refresh));
    refresh();
  }

  // ---------------- step 3: blank / zero status tags ----------------
  if (page === "history") {
    function status(inp) {
      const cell = inp.parentElement.nextElementSibling;
      const v = inp.value.replace(/[,\s$]/g, "");
      if (v === "") { cell.innerHTML = '<span class="tag unk">Unknown</span>'; inp.classList.remove("z"); }
      else if (isNaN(Number(v))) { cell.innerHTML = '<span class="tag bad">Not a number</span>'; }
      else if (Number(v) === 0) { cell.innerHTML = '<span class="tag zer">Confirmed zero</span>'; inp.classList.add("z"); }
      else { cell.innerHTML = '<span class="tag ok">Entered</span>'; inp.classList.remove("z"); }
    }
    function firstGr() {
      let first = null;
      $$('input.hist[name^="gr_"]').forEach(i => {
        const y = parseInt(i.name.slice(3), 10), v = num(i.value);
        if (v > 0 && (first === null || y < first)) first = y;
      });
      const b = $("#firstgr b"); if (b) b.textContent = first === null ? "none yet" : first;
    }
    $$("input.hist").forEach(i => { status(i); i.addEventListener("input", () => { status(i); firstGr(); }); });
    firstGr();
    $$('input[name="pte"]').forEach(r => r.addEventListener("change", () => {
      $("#pteyears").style.display = (val("pte") === "yes") ? "" : "none";
    }));
  }

  // ---------------- step 4: tabs + live totals ----------------
  if (page === "expenses") {
    $$(".tab").forEach(b => b.addEventListener("click", () => {
      $$(".tab").forEach(x => x.classList.toggle("active", x === b));
      $$(".pane").forEach(p => p.hidden = (p.id !== "p-" + b.dataset.t));
      $("#tabfield").value = b.dataset.t;
    }));
    function qualified(L) {
      const hasPct = !L.dataset.nopct;
      let sum = 0;
      $$("tbody tr", L).forEach(tr => {
        const a = $(".amt", tr).value, p = hasPct ? $(".pct", tr).value : (a === "" ? "" : "100");
        const cell = $(".calc", tr);
        if (a === "" && p === "") { cell.textContent = "—"; return; }
        const q = num(a) * num(p) / 100; sum += q; cell.textContent = money(q);
      });
      $(".lt", L).textContent = money(sum);
      const ta = $(".tamt", L).value, tp = hasPct ? $(".tpct", L).value : (ta === "" ? "" : "100");
      const totalMode = (ta !== "" || (hasPct && tp !== ""));
      const tq = num(ta) * num(tp) / 100;
      $(".tres", L).textContent = totalMode ? money(tq) : "—";
      $("tbody", L).classList.toggle("muted", totalMode);
      $(".lt-row", L).classList.toggle("muted", totalMode);
      const used = totalMode ? tq : sum;
      $(".used", L).textContent = money(used);
      return used;
    }
    function calc() {
      $$(".lines").forEach(L => {
        const q = qualified(L), cat = L.dataset.cat;
        if (cat === "contract") { $("#c_q").textContent = money(q); $("#c_65").textContent = money(q * 0.65); }
        if (cat === "foreign") { $("#f_amt").textContent = money(q); $("#f_am").textContent = money(q / 30); $("#f_red").textContent = money(q - q / 30); }
      });
    }
    $$(".lines").forEach(L => L.addEventListener("input", calc));
    calc();
    const wb = $("#wbfile");
    if (wb) wb.addEventListener("change", () => { if (wb.files.length) $("#stepform").submit(); });
  }

  // ---------------- step 6: busy indicator ----------------
  if (page === "review") {
    const f = $("#runform");
    if (f) f.addEventListener("submit", () => { $("#runbtn").disabled = true; $("#working").classList.remove("hidden"); });
  }
})();

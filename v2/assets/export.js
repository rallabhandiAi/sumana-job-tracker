/* Excel export for the tracker: builds a 3-tab workbook (Dashboard, Jobs, Search Profile) with ExcelJS. */
"use strict";
const STATUSES = ["New", "Shortlisted", "Applied", "Recruiter screen", "Interviewing", "Offer", "Rejected", "Not a fit", "Closed"];

function isoToUTCDate(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso || ""));
  return m ? new Date(Date.UTC(+m[1], +m[2] - 1, +m[3])) : null;
}

// rows: merged role objects (job fields + status/applied_on/contact/notes), already in display order.
// profile: { title, subtitle, rows: [[label, text], ...] } read from data/profile.json in the private data repo.
function buildWorkbook(ExcelJS, rows, todayISO, profile, contacts) {
  profile = profile || {};
  contacts = Array.isArray(contacts) ? contacts : [];
  const MAX_ROW = 400;
  const FONT = "Arial";
  const DUSK = "FF2E2A4F", MARIGOLD = "FFE3A21A", GRID = "FFD9D6E6", INK = "FF1F1D33", MUTED = "FF5B5873";
  const thin = { style: "thin", color: { argb: GRID } };
  const border = { top: thin, left: thin, bottom: thin, right: thin };
  const font = (o = {}) => Object.assign({ name: FONT, size: 10, color: { argb: INK } }, o);
  const fill = (argb) => ({ type: "pattern", pattern: "solid", fgColor: { argb } });
  const today = isoToUTCDate(todayISO);

  const wb = new ExcelJS.Workbook();
  wb.creator = "Claude";
  wb.calcProperties.fullCalcOnLoad = true;

  // ---------------- Dashboard (first tab)
  const db = wb.addWorksheet("Dashboard", { views: [{ showGridLines: false }] });
  // ---------------- Jobs
  const ws = wb.addWorksheet("Jobs", { views: [{ state: "frozen", xSplit: 4, ySplit: 1 }] });
  const cols = [
    ["#", 5, false], ["Match (1-5)", 9, false], ["Status", 16, true], ["Next Action", 26, true], ["Job Title", 42, false],
    ["Company", 26, false], ["Location", 30, false], ["Work Mode", 11, false], ["Type", 20, false],
    ["Pay", 16, false], ["Skill Tags", 22, false], ["Posted", 18, false], ["Found On", 12, false],
    ["Source", 15, false], ["Link", 11, false], ["Why It Fits", 46, false], ["Watch-outs", 44, false],
    ["Work Authorization", 24, false], ["Applied On", 12, true], ["Follow-up By", 12, false],
    ["Recruiter / Contact", 24, true], ["Notes", 40, true],
  ];
  const L = {};
  cols.forEach(([name, width, editable], i) => {
    const letter = ws.getColumn(i + 1).letter;
    L[name] = letter;
    ws.getColumn(i + 1).width = width;
    const c = ws.getCell(`${letter}1`);
    c.value = name;
    c.font = font({ bold: true, color: { argb: editable ? INK : "FFFFFFFF" } });
    c.fill = fill(editable ? MARIGOLD : DUSK);
    c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    c.border = border;
  });
  ws.getRow(1).height = 30;
  const wrapW = { "Next Action": 26, "Job Title": 42, "Company": 26, "Location": 30, "Type": 20, "Pay": 16, "Skill Tags": 22, "Posted": 18, "Why It Fits": 46, "Watch-outs": 44, "Work Authorization": 24, "Notes": 40 };
  const centered = new Set(["#", "Match (1-5)", "Work Mode", "Found On", "Link", "Applied On", "Follow-up By"]);
  const lines = (t, w) => Math.max(1, Math.ceil(String(t || "").length / (w * 1.05)));
  const last = rows.length + 1;

  rows.forEach((j, i) => {
    const r = i + 2;
    const applied = isoToUTCDate(j.applied_on);
    const vals = {
      "#": i + 1, "Match (1-5)": j.fit || null, "Status": j.status || "New", "Next Action": j.next_action || "", "Job Title": j.title || "",
      "Company": j.company || "", "Location": j.location || "", "Work Mode": j.mode || "",
      "Type": j.type || "", "Pay": j.pay || "", "Skill Tags": (j.tags || []).join(", "),
      "Posted": j.posted || "", "Found On": isoToUTCDate(j.found_on), "Source": j.source || "",
      "Link": /^https?:\/\//i.test(j.url || "") ? { text: "Open posting", hyperlink: j.url } : "",
      "Why It Fits": j.why || "", "Watch-outs": j.watch || "", "Work Authorization": j.auth || "",
      "Applied On": applied, "Recruiter / Contact": j.contact || "", "Notes": j.notes || "",
    };
    for (const [name, v] of Object.entries(vals)) {
      const c = ws.getCell(`${L[name]}${r}`);
      c.value = v;
      c.font = font();
      c.border = border;
      c.alignment = { vertical: "top", wrapText: name in wrapW, horizontal: centered.has(name) ? "center" : undefined };
    }
    ws.getCell(`${L["Found On"]}${r}`).numFmt = "mmm d, yyyy";
    ws.getCell(`${L["Match (1-5)"]}${r}`).font = font({ bold: true, size: 11 });
    ws.getCell(`${L["Job Title"]}${r}`).font = font({ bold: true });
    ws.getCell(`${L["Link"]}${r}`).font = font({ color: { argb: "FF1F5FBF" }, underline: true });
    for (const name of ["Status", "Next Action", "Applied On", "Recruiter / Contact", "Notes"]) {
      ws.getCell(`${L[name]}${r}`).fill = fill("FFFFF8E6");
    }
    const h = Math.max(...Object.entries(wrapW).map(([n, w]) => lines(vals[n] && vals[n].text ? vals[n].text : vals[n], w)));
    ws.getRow(r).height = Math.max(28, 13.5 * h + 6);
  });

  for (let r = 2; r <= MAX_ROW; r++) {
    const a = ws.getCell(`${L["Applied On"]}${r}`);
    a.numFmt = "mmm d, yyyy";
    const fu = ws.getCell(`${L["Follow-up By"]}${r}`);
    const av = a.value instanceof Date ? new Date(a.value.getTime() + 7 * 864e5) : "";
    const explicit = r <= last ? isoToUTCDate(rows[r - 2].follow_up_on) : null;
    fu.value = explicit || { formula: `IF(${L["Applied On"]}${r}="","",${L["Applied On"]}${r}+7)`, result: av };
    fu.numFmt = "mmm d, yyyy";
    if (r > last) {
      for (const name of ["Status", "Next Action", "Applied On", "Follow-up By", "Recruiter / Contact", "Notes"]) {
        const c = ws.getCell(`${L[name]}${r}`);
        c.border = border;
        c.font = font();
      }
    } else {
      fu.border = border;
      fu.font = font();
      fu.alignment = { vertical: "top", horizontal: "center" };
    }
  }

  ws.dataValidations.add(`${L.Status}2:${L.Status}${MAX_ROW}`, { type: "list", allowBlank: true, formulae: [`"${STATUSES.join(",")}"`], showErrorMessage: true, errorTitle: "Status", error: "Pick a status from the list." });
  ws.dataValidations.add(`${L["Work Mode"]}2:${L["Work Mode"]}${MAX_ROW}`, { type: "list", allowBlank: true, formulae: ['"Remote,Hybrid,On-site,Confirm"'] });
  ws.dataValidations.add(`${L["Match (1-5)"]}2:${L["Match (1-5)"]}${MAX_ROW}`, { type: "whole", operator: "between", allowBlank: true, formulae: [1, 5], showErrorMessage: true, error: "Match is a whole number from 1 to 5." });

  const st = `$${L.Status}2`, fu1 = `$${L["Follow-up By"]}2`;
  ws.addConditionalFormatting({ ref: `A2:${L.Notes}${MAX_ROW}`, rules: [{ type: "expression", priority: 1, formulae: [`OR(${st}="Rejected",${st}="Not a fit",${st}="Closed")`], style: { font: { color: { argb: "FF9A97AB" }, italic: true } } }] });
  ws.addConditionalFormatting({ ref: `${L["Match (1-5)"]}2:${L["Match (1-5)"]}${MAX_ROW}`, rules: [
    { type: "cellIs", priority: 2, operator: "equal", formulae: ["5"], style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFF6C453" } } } },
    { type: "cellIs", priority: 3, operator: "equal", formulae: ["4"], style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFFCEBC0" } } } },
    { type: "cellIs", priority: 4, operator: "equal", formulae: ["3"], style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFEEEDF5" } } } },
  ] });
  const statusFills = [["Shortlisted", "FFEFE6FA", false], ["Applied", "FFDCEBFB", false], ["Recruiter screen", "FFE4DDF8", false], ["Interviewing", "FFD9CCF5", true], ["Offer", "FFCDEFD6", true]];
  ws.addConditionalFormatting({ ref: `${L.Status}2:${L.Status}${MAX_ROW}`, rules: statusFills.map(([v, argb, bold], k) => ({ type: "cellIs", priority: 5 + k, operator: "equal", formulae: [`"${v}"`], style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb } }, font: { bold } } })) });
  ws.addConditionalFormatting({ ref: `${L["Follow-up By"]}2:${L["Follow-up By"]}${MAX_ROW}`, rules: [{ type: "expression", priority: 12, formulae: [`AND(${fu1}<>"",${fu1}<=TODAY(),OR(${st}="Applied",${st}="Recruiter screen"))`], style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFD64545" } }, font: { bold: true, color: { argb: "FFFFFFFF" } } } }] });
  ws.autoFilter = `A1:${L.Notes}${Math.max(2, last)}`;

  // ---------------- Dashboard content (formulas mirror the Jobs tab; results pre-computed for previewers)
  [26, 12, 4, 30, 12, 4, 24, 40].forEach((w, i) => { db.getColumn(i + 1).width = w; });
  const set = (addr, value, f = {}, extra = {}) => { const c = db.getCell(addr); c.value = value; c.font = font(f); Object.assign(c, extra); return c; };
  set("A1", profile.title || "Job Search", { size: 18, bold: true, color: { argb: DUSK } });
  set("A2", profile.subtitle || "", { color: { argb: MUTED } });
  set("A3", `Exported from the live tracker on ${todayISO}. New roles are added at 10:30 AM and 8:30 PM CT.`, { size: 9, italic: true, color: { argb: MUTED } });
  const header = (addr, text, spanCols) => {
    set(addr, text, { size: 11, bold: true, color: { argb: "FFFFFFFF" } }, { fill: fill(DUSK) });
    (spanCols || []).forEach((col) => { db.getCell(`${col}${addr.replace(/\D/g, "")}`).fill = fill(DUSK); });
  };
  const rng = (name) => `Jobs!$${L[name]}$2:$${L[name]}$${MAX_ROW}`;
  const statusOf = rows.map((j) => j.status || "New");
  const cnt = (s) => statusOf.filter((x) => x === s).length;
  const lineRow = (r, lc, vc, label, formula, result, fmt, bold) => {
    set(`${lc}${r}`, label, { bold: !!bold }, { border: { bottom: thin } });
    const c = set(`${vc}${r}`, typeof formula === "number" ? formula : { formula, result }, { bold: true }, { border: { bottom: thin }, alignment: { horizontal: "right" } });
    if (fmt) c.numFmt = fmt;
    return c;
  };
  header("A5", "Pipeline", ["B"]);
  STATUSES.forEach((s, i) => lineRow(6 + i, "A", "B", s, `COUNTIF(${rng("Status")},"${s}")`, cnt(s)));
  lineRow(15, "A", "B", "Total roles tracked", `COUNTA(${rng("Job Title")})`, rows.length, null, true);

  const appliedDates = rows.map((j) => isoToUTCDate(j.applied_on)).filter(Boolean);
  const weekAgo = new Date(today.getTime() - 6 * 864e5);
  const applied7 = appliedDates.filter((d) => d >= weekAgo && d <= today).length;
  const due = rows.filter((j) => { const d = isoToUTCDate(j.applied_on); return d && (j.status === "Applied" || j.status === "Recruiter screen") && d.getTime() + 7 * 864e5 <= today.getTime(); }).length;
  const resp = cnt("Recruiter screen") + cnt("Interviewing") + cnt("Offer");
  header("D5", "This week", ["E"]);
  const goal = lineRow(6, "D", "E", "Weekly application goal", 8);
  goal.font = font({ bold: true, color: { argb: "FF0000FF" } });
  goal.fill = fill("FFFFF2B3");
  goal.note = "Your input: change the weekly goal here.";
  lineRow(7, "D", "E", "Applied in the last 7 days", `COUNTIFS(${rng("Applied On")},">="&(TODAY()-6),${rng("Applied On")},"<="&TODAY())`, applied7);
  lineRow(8, "D", "E", "Progress to goal", "IFERROR(E7/E6,0)", applied7 / 8, "0%");
  lineRow(9, "D", "E", "Follow-ups due now", `COUNTIFS(${rng("Follow-up By")},"<="&TODAY(),${rng("Status")},"Applied")+COUNTIFS(${rng("Follow-up By")},"<="&TODAY(),${rng("Status")},"Recruiter screen")`, due);
  lineRow(10, "D", "E", "Applications sent (all time)", `COUNT(${rng("Applied On")})`, appliedDates.length);
  lineRow(11, "D", "E", "Response rate", "IFERROR((B9+B10+B11)/E10,0)", appliedDates.length ? resp / appliedDates.length : 0, "0%");
  lineRow(12, "D", "E", "Interview rate", "IFERROR((B10+B11)/E10,0)", appliedDates.length ? (cnt("Interviewing") + cnt("Offer")) / appliedDates.length : 0, "0%");
  set("D13", "Response = reached recruiter screen or later. Rates use roles with an Applied On date.", { size: 8, italic: true, color: { argb: MUTED } });

  header("G5", "Role mix", ["H"]);
  const modes = rows.map((j) => j.mode || ""), types = rows.map((j) => String(j.type || ""));
  const mix = [
    ["Match 5 – apply first", `COUNTIF(${rng("Match (1-5)")},5)`, rows.filter((j) => j.fit === 5).length],
    ["Match 4", `COUNTIF(${rng("Match (1-5)")},4)`, rows.filter((j) => j.fit === 4).length],
    ["Match 3", `COUNTIF(${rng("Match (1-5)")},3)`, rows.filter((j) => j.fit === 3).length],
    ["Remote", `COUNTIF(${rng("Work Mode")},"Remote")`, modes.filter((m) => m === "Remote").length],
    ["Hybrid / On-site (Chicago)", `COUNTIF(${rng("Work Mode")},"Hybrid")+COUNTIF(${rng("Work Mode")},"On-site")`, modes.filter((m) => m === "Hybrid" || m === "On-site").length],
    ["Remote policy to confirm", `COUNTIF(${rng("Work Mode")},"Confirm")`, modes.filter((m) => m === "Confirm").length],
    ["Contract", `COUNTIF(${rng("Type")},"Contract*")`, types.filter((t) => t.startsWith("Contract")).length],
    ["Full-time", `COUNTIF(${rng("Type")},"Full-time*")`, types.filter((t) => t.startsWith("Full-time")).length],
  ];
  mix.forEach(([label, f, res], i) => { const c = lineRow(6 + i, "G", "H", label, f, res); c.alignment = { horizontal: "left" }; });

  header("A18", "How to use this sheet", ["B", "C", "D", "E", "F", "G", "H"]);
  [
    "On the Jobs tab, edit only the marigold columns: Status, Next Action, Applied On, Recruiter / Contact, Notes.",
    "Follow-up By fills itself (Applied On + 7 days, unless a date was set in the tracker) and turns red when due.",
    "Closed, Rejected and Not a fit rows grey out automatically. Use the filter arrows to sort by Match or Status.",
    "Several vendor postings are the same end-client job. Submit through one vendor only; duplicates can get you disqualified.",
    "This is a snapshot. Update statuses in the live tracker so the twice-daily search and your Excel exports stay in sync.",
  ].forEach((t, i) => set(`A${19 + i}`, `•  ${t}`));
  set("A25", "Example of a filled-in row (format only):", { bold: true });
  const exCols = ["A", "B", "D", "E", "G"], exHead = ["Status", "Applied On", "Follow-up By", "Recruiter / Contact", "Notes"];
  const exDate = new Date(Date.UTC(2026, 9, 6));
  const exVals = ["Applied", exDate, { formula: "B27+7", result: new Date(exDate.getTime() + 7 * 864e5) }, "Recruiter name – vendor", "Sent tailored resume v2 (low-power focus); asked W2 rate"];
  exCols.forEach((col, i) => {
    set(`${col}26`, exHead[i], { size: 9, bold: true }, { fill: fill(MARIGOLD) });
    set(`${col}27`, exVals[i], { size: 9, italic: true, color: { argb: MUTED } }, { fill: fill("FFFFF8E6") });
  });
  db.getCell("B27").numFmt = "mmm d, yyyy";
  db.getCell("D27").numFmt = "mmm d, yyyy";
  set("A29", "Every posting was opened and checked as live when it was added (Dice, ZipRecruiter, Built In, Glassdoor, The Ladders, company career pages).", { size: 8, italic: true, color: { argb: MUTED } });

  // ---------------- Contacts
  const cs = wb.addWorksheet("Contacts", { views: [{ state: "frozen", ySplit: 1 }] });
  const ccols = [["Name", 22], ["Company", 22], ["Relationship", 16], ["Related Role", 36], ["How To Reach", 28], ["Last Contact", 13], ["Next Follow-up", 14], ["Status", 11], ["Notes", 50]];
  ccols.forEach(([name, width], i) => {
    cs.getColumn(i + 1).width = width;
    const c = cs.getCell(1, i + 1);
    c.value = name; c.font = font({ bold: true, color: { argb: "FFFFFFFF" } }); c.fill = fill(DUSK);
    c.alignment = { horizontal: "center", vertical: "middle", wrapText: true }; c.border = border;
  });
  cs.getRow(1).height = 24;
  contacts.forEach((ct, i) => {
    const r = i + 2;
    const vals = [ct.name || "", ct.company || "", ct.relationship || "", ct.related_role || "", ct.reach || "",
      isoToUTCDate(ct.last_contact), isoToUTCDate(ct.next_follow_up), ct.status || "", ct.notes || ""];
    vals.forEach((v, k) => {
      const c = cs.getCell(r, k + 1);
      c.value = v; c.font = font(); c.border = border; c.alignment = { vertical: "top", wrapText: k === 3 || k === 8 };
      if (k === 5 || k === 6) c.numFmt = "mmm d, yyyy";
    });
    const due = ct.status !== "Done" && ct.next_follow_up && isoToUTCDate(ct.next_follow_up) <= today;
    if (due) cs.getCell(r, 7).fill = fill("FFFBE3E3");
  });
  if (!contacts.length) { cs.getCell("A2").value = "No contacts yet."; cs.getCell("A2").font = font({ italic: true, color: { argb: MUTED } }); }

  // ---------------- Search Profile
  const sp = wb.addWorksheet("Search Profile", { views: [{ showGridLines: false }] });
  sp.getColumn(1).width = 28;
  sp.getColumn(2).width = 110;
  sp.getCell("A1").value = "Search profile used for curation and the twice-daily search";
  sp.getCell("A1").font = font({ size: 14, bold: true, color: { argb: DUSK } });
  const profileRows = Array.isArray(profile.rows) && profile.rows.length ? profile.rows : [["Search profile", "Not loaded: it lives in data/profile.json in the private data repository."]];
  profileRows.forEach(([k, v], i) => {
    const r = i + 3;
    const a = sp.getCell(`A${r}`), b = sp.getCell(`B${r}`);
    a.value = k; a.font = font({ bold: true, color: { argb: DUSK } }); a.alignment = { vertical: "top" };
    b.value = String(v); b.font = font(); b.alignment = { vertical: "top", wrapText: true };
    sp.getRow(r).height = 13.5 * Math.max(1, Math.ceil(String(v).length / 112)) + 6;
  });
  return wb;
}

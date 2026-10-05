/* Sumana's Job Tracker: static GitHub Pages app.
   Data lives in the repo as JSON (data/jobs.json, data/tracking.json, data/meta.json).
   Reads: same-site static files, or the GitHub Contents API when a token is set (needed for private repos).
   Writes: GitHub Contents API with a fine-grained token stored in this browser. */
"use strict";
(() => {
  const $ = (s) => document.querySelector(s);
  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const CT = "America/Chicago";
  const CLOSED = new Set(["Rejected", "Not a fit", "Closed"]);
  const STAGES = [["New", "New"], ["Shortlisted", "Shortlisted"], ["Applied", "Applied"], ["Recruiter screen", "Screen"], ["Interviewing", "Interviewing"], ["Offer", "Offer"]];
  const FILES = { jobs: "data/jobs.json", tracking: "data/tracking.json", meta: "data/meta.json", profile: "data/profile.json" };
  const LS_KEY = "sjt.settings.v1";
  const EXCELJS_URL = "https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js";
  const EXCELJS_SRI = "sha384-Pqp51FUN2/qzfxZxBCtF0stpc9ONI6MYZpVqmo8m20SoaQCzf+arZvACkLkirlPz";

  // ---------------------------------------------------------------- settings
  function detectRepo() {
    const host = /^([a-z0-9-]+)\.github\.io$/i.exec(location.hostname);
    if (!host) return { owner: "", repo: "" };
    const first = location.pathname.split("/").filter(Boolean)[0];
    return { owner: host[1], repo: first && !/\.html?$/i.test(first) ? first : `${host[1]}.github.io` };
  }
  function loadSettings() {
    let saved = {};
    try { saved = JSON.parse(localStorage.getItem(LS_KEY) || "{}") || {}; } catch (_) { saved = {}; }
    const c = window.TRACKER_CONFIG || {};
    const d = detectRepo();
    return {
      owner: saved.owner || c.owner || d.owner,
      repo: saved.repo || c.dataRepo || d.repo,
      branch: saved.branch || c.branch || "main",
      token: saved.token || "",
    };
  }
  function storeSettings(s) {
    try { localStorage.setItem(LS_KEY, JSON.stringify(s)); return true; } catch (_) { return false; }
  }
  let cfg = loadSettings();
  const canUseApi = () => Boolean(cfg.token && cfg.owner && cfg.repo);
  const sameSiteRepo = () => { const d = detectRepo(); return d.owner && d.owner.toLowerCase() === (cfg.owner || "").toLowerCase() && d.repo.toLowerCase() === (cfg.repo || "").toLowerCase(); };

  // ---------------------------------------------------------------- GitHub data layer
  const API = "https://api.github.com";
  const shas = {};
  const b64enc = (str) => { const bytes = new TextEncoder().encode(str); let bin = ""; for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000)); return btoa(bin); };
  const b64dec = (b64) => { const bin = atob(String(b64).replace(/\s/g, "")); const bytes = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i); return new TextDecoder().decode(bytes); };
  const headers = (accept) => ({ Accept: accept || "application/vnd.github+json", "X-GitHub-Api-Version": "2022-11-28", Authorization: `Bearer ${cfg.token}` });
  const contentsUrl = (path) => `${API}/repos/${encodeURIComponent(cfg.owner)}/${encodeURIComponent(cfg.repo)}/contents/${path}`;
  function httpError(status, message) { const e = new Error(message || `HTTP ${status}`); e.status = status; return e; }

  async function readJSON(key) {
    const path = FILES[key];
    if (canUseApi()) {
      const url = `${contentsUrl(path)}?ref=${encodeURIComponent(cfg.branch)}`;
      const r = await fetch(url, { headers: headers(), cache: "no-store" });
      if (r.status === 404) { shas[key] = null; return null; }
      if (!r.ok) throw httpError(r.status);
      const meta = await r.json();
      shas[key] = meta.sha;
      let text;
      if (meta.encoding === "base64" && meta.content) text = b64dec(meta.content);
      else {
        const raw = await fetch(url, { headers: headers("application/vnd.github.raw+json"), cache: "no-store" });
        if (!raw.ok) throw httpError(raw.status);
        text = await raw.text();
      }
      return JSON.parse(text);
    }
    if (location.protocol === "file:") throw httpError(0, "file");
    // A separate (private) data repo can only be read through the API with a token.
    if (cfg.owner && cfg.repo && detectRepo().owner && !sameSiteRepo()) throw httpError(404);
    const r = await fetch(`${path}?v=${Date.now()}`, { cache: "no-store" });
    if (r.status === 404) return null;
    if (!r.ok) throw httpError(r.status);
    return r.json();
  }

  async function writeJSON(key, obj, message) {
    const body = { message, content: b64enc(JSON.stringify(obj, null, 2) + "\n"), branch: cfg.branch };
    if (shas[key]) body.sha = shas[key];
    const r = await fetch(contentsUrl(FILES[key]), { method: "PUT", headers: Object.assign(headers(), { "Content-Type": "application/json" }), body: JSON.stringify(body) });
    if (!r.ok) throw httpError(r.status);
    const out = await r.json();
    shas[key] = out && out.content ? out.content.sha : null;
  }

  // ---------------------------------------------------------------- state
  let pendingWrites = 0;
  const state = {
    jobs: [], tracking: {}, meta: null, profile: null, loaded: false, loadError: null, writable: false,
    f: { q: "", where: "all", type: "all", match: 0, showClosed: false, sort: "match", stage: null },
    open: new Set(), pendingRender: false,
  };

  // ---------------------------------------------------------------- dates
  const pad = (n) => String(n).padStart(2, "0");
  const localISO = (d = new Date()) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  const addDaysISO = (iso, n) => { const [y, m, d] = iso.split("-").map(Number); return localISO(new Date(y, m - 1, d + n)); };
  const shortDate = (iso) => { const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso || ""); if (!m) return iso || ""; return new Date(+m[1], +m[2] - 1, +m[3]).toLocaleDateString("en-US", { month: "short", day: "numeric" }); };
  const fmtCT = (ts) => { const d = new Date(ts); return isNaN(d) ? "" : d.toLocaleString("en-US", { timeZone: CT, weekday: "short", month: "short", day: "numeric", hour: "numeric", minute: "2-digit" }); };
  function nextRunText() {
    const p = Object.fromEntries(new Intl.DateTimeFormat("en-US", { timeZone: CT, hour: "numeric", minute: "numeric", hourCycle: "h23" }).formatToParts(new Date()).map((x) => [x.type, x.value]));
    const mins = (+p.hour) * 60 + (+p.minute);
    if (mins < 630) return "Next run: today 10:30 AM";
    if (mins < 1230) return "Next run: today 8:30 PM";
    return "Next run: tomorrow 10:30 AM";
  }
  function postedLabel(p) {
    const raw = String(p || "").trim();
    if (!raw || /^not (shown|stated|listed)$/i.test(raw)) return "";
    const s = raw.replace(/\d{4}-\d{2}-\d{2}/g, (iso) => shortDate(iso));
    if (/^\d+\+? days?$/i.test(s)) return `Posted ${s} ago`;
    return /^[~\d]|^[A-Z][a-z]{2} \d/.test(s) ? `Posted ${s}` : s;
  }

  // ---------------------------------------------------------------- derived rows
  function merged() {
    return state.jobs.filter((j) => j && j.id).map((j) => {
      const t = state.tracking[j.id] || {};
      return Object.assign({}, j, { status: t.status || "New", applied_on: t.applied_on || "", contact: t.contact || "", notes: t.notes || "" });
    });
  }
  const isChicago = (j) => j.mode === "Hybrid" || j.mode === "On-site" || /chicago|naperville|illinois|,\s*IL\b|schaumburg|rolling meadows|batavia|lemont|itasca/i.test(j.location || "");
  const isRemote = (j) => j.mode === "Remote" || j.mode === "Confirm";
  const isPassed = (j) => CLOSED.has(j.status) || j.posting_status === "Closed";
  const followDue = (j, today) => j.applied_on && (j.status === "Applied" || j.status === "Recruiter screen") && addDaysISO(j.applied_on, 7) <= today;
  function sortRows(rows, how) {
    const byMatch = (a, b) => (b.fit || 0) - (a.fit || 0) || String(b.found_on || "").localeCompare(String(a.found_on || "")) || (a.rank ?? 9999) - (b.rank ?? 9999);
    if (how === "newest") return rows.sort((a, b) => String(b.found_on || "").localeCompare(String(a.found_on || "")) || String(b.found_run || "").localeCompare(String(a.found_run || "")) || byMatch(a, b));
    if (how === "company") return rows.sort((a, b) => String(a.company || "").localeCompare(String(b.company || "")) || byMatch(a, b));
    return rows.sort(byMatch);
  }
  function visibleRows() {
    const f = state.f, q = f.q.trim().toLowerCase();
    const rows = merged().filter((j) => {
      if (f.stage) { if (j.status !== f.stage) return false; }
      else if (!f.showClosed && isPassed(j)) return false;
      if (f.where === "remote" && !isRemote(j)) return false;
      if (f.where === "chicago" && !isChicago(j)) return false;
      if (f.type === "contract" && !/^contract/i.test(j.type || "")) return false;
      if (f.type === "fulltime" && !/^full-time/i.test(j.type || "")) return false;
      if (f.match && (j.fit || 0) < f.match) return false;
      if (q && ![j.title, j.company, j.location, j.type, j.why, (j.tags || []).join(" ")].join(" ").toLowerCase().includes(q)) return false;
      return true;
    });
    return sortRows(rows, f.sort);
  }

  // ---------------------------------------------------------------- rendering
  const pips = (n) => Array.from({ length: 5 }, (_, i) => `<span class="pip${i < (n || 0) ? " on" : ""}"></span>`).join("");
  function roleHTML(j, today) {
    const id = esc(j.id);
    const href = /^https?:\/\//i.test(j.url || "") ? esc(j.url) : "";
    const isNew = j.found_on && j.found_on >= addDaysISO(today, -1) && j.status === "New";
    const due = followDue(j, today);
    const dis = state.writable ? "" : " disabled";
    const opts = STATUSES.map((s) => `<option${s === j.status ? " selected" : ""}>${esc(s)}</option>`).join("");
    const pl = postedLabel(j.posted);
    const chips = [
      j.mode ? `<span class="chip mode-${esc(String(j.mode).replace(/\s+/g, "-"))}">${esc(j.mode === "Confirm" ? "Remote? confirm" : j.mode)}</span>` : "",
      j.type ? `<span class="chip">${esc(j.type)}</span>` : "",
      j.pay && j.pay !== "Not listed" ? `<span class="chip">${esc(j.pay)}</span>` : "",
      pl ? `<span class="chip quiet">${esc(pl)}</span>` : "",
      j.found_on ? `<span class="chip quiet">Found ${esc(shortDate(j.found_on))}</span>` : "",
    ].join("");
    const tags = (Array.isArray(j.tags) ? j.tags : []).map((t) => `<span class="tag">${esc(t)}</span>`).join("");
    const track = j.applied_on ? `<div class="trackline">Applied ${esc(shortDate(j.applied_on))} · follow up by ${esc(shortDate(addDaysISO(j.applied_on, 7)))}</div>` : "";
    return `<article class="role${isPassed(j) ? " is-dim" : ""}" data-id="${id}">
      <div class="match" aria-label="Match ${esc(j.fit || "not scored")} of 5"><span class="pips" aria-hidden="true">${pips(j.fit)}</span><span class="mlabel">Match <b>${esc(j.fit || "–")}</b></span></div>
      <div class="rbody">
        <div class="titleline"><h3>${esc(j.title)}</h3>${isNew ? '<span class="badge badge-new">New</span>' : ""}${j.posting_status === "Closed" ? '<span class="badge badge-closed">Posting closed</span>' : ""}${due ? '<span class="badge badge-due">Follow up</span>' : ""}${j.added_by === "manual" ? '<span class="badge badge-mine">Added by you</span>' : ""}</div>
        <div class="org"><b>${esc(j.company)}</b> · ${esc(j.location)}</div>
        <div class="chips">${chips}</div>
        ${tags ? `<div class="tags" aria-label="Skill match">${tags}</div>` : ""}
      </div>
      <div class="side">
        <label class="sr-only" for="st-${id}">Status for ${esc(j.title)}</label>
        <select class="status" id="st-${id}" data-s="${esc(j.status)}" data-act="status"${dis}>${opts}</select>
        ${href ? `<a class="open" href="${href}" target="_blank" rel="noopener noreferrer">Open posting ↗</a>` : ""}
        ${track}
      </div>
      <details class="more" data-id="${id}"${state.open.has(j.id) ? " open" : ""}>
        <summary>Why it fits, watch-outs and your notes</summary>
        <div class="more-grid">
          <div><h4>Why it fits</h4><p>${esc(j.why || "—")}</p></div>
          <div><h4>Watch-outs</h4><p>${esc(j.watch || "—")}</p></div>
          <div><h4>Work authorization</h4><p>${esc(j.auth || "Not stated")}</p></div>
          <div><h4>Source</h4><p>${esc(j.source || "—")}${j.last_checked ? ` · checked ${esc(shortDate(j.last_checked))}` : ""}</p></div>
          <div class="track-grid">
            <div class="field"><label for="ap-${id}">Applied on</label><input type="date" id="ap-${id}" data-act="applied_on" value="${esc(j.applied_on)}"${dis}></div>
            <div class="field"><label for="ct-${id}">Recruiter / contact</label><input id="ct-${id}" data-act="contact" value="${esc(j.contact)}" placeholder="Name, vendor, phone or email"${dis}></div>
            <div class="field full"><label for="nt-${id}">Notes</label><textarea id="nt-${id}" data-act="notes" placeholder="Resume version sent, rate discussed, interview dates…"${dis}>${esc(j.notes)}</textarea></div>
          </div>
        </div>
      </details>
    </article>`;
  }

  function renderPipeline(all, today) {
    const counts = {};
    all.forEach((j) => { counts[j.status] = (counts[j.status] || 0) + 1; });
    $("#pipeline").innerHTML = STAGES.map(([s, label]) => `<button type="button" class="stage" data-stage="${esc(s)}" aria-pressed="${state.f.stage === s}"><span class="n">${counts[s] || 0}</span><span class="l">${esc(label)}</span></button>`).join("");
    const due = all.filter((j) => followDue(j, today)).length;
    const applied = all.filter((j) => j.applied_on).length;
    const responded = all.filter((j) => j.applied_on && ["Recruiter screen", "Interviewing", "Offer"].includes(j.status)).length;
    const bits = [];
    if (due) bits.push(`<span class="pill pill-bad">${due} follow-up${due > 1 ? "s" : ""} due</span>`);
    if (applied) bits.push(`<span>${applied} applied · ${Math.round((responded / applied) * 100)}% response rate</span>`);
    bits.push(`<span>${all.filter(isPassed).length} closed or passed</span>`);
    if (state.f.stage) bits.push(`<button type="button" class="linkish" data-act="clearstage">Show all statuses</button>`);
    $("#pipenotes").innerHTML = bits.join("");
  }

  function renderRunline() {
    const m = state.meta;
    $("#lastrun").innerHTML = m && m.last_run_at ? `Last run: <strong>${esc(fmtCT(m.last_run_at))}</strong>${typeof m.last_new_count === "number" ? ` · ${m.last_new_count} new` : ""}` : "";
    $("#nextrun").textContent = nextRunText();
    const runs = m && Array.isArray(m.runs) ? m.runs.slice(0, 14) : [];
    $("#logBody").innerHTML = runs.length
      ? runs.map((r) => `<tr><td>${esc(r.label || fmtCT(r.at))}</td><td>${esc(r.new ?? "–")}</td><td>${esc(r.checked ?? "–")}</td><td>${esc(r.closed ?? "–")}</td></tr>`).join("")
      : `<tr><td colspan="4">No runs recorded yet.</td></tr>`;
  }

  function renderNotice() {
    const n = $("#notice");
    let text = "";
    if (state.loadError === "file") text = "Open this tracker through GitHub Pages or a local web server (see README); browsers block data files opened straight from disk.";
    else if (state.loadError === "auth") text = "GitHub didn't accept the saved token. Check it in Settings.";
    else if (state.loadError === "notfound") text = canUseApi()
      ? "No data files found in that repository. Check the owner, repository and branch in Settings."
      : (cfg.owner && cfg.repo ? `Add your GitHub token in Settings to load the roles from ${cfg.owner}/${cfg.repo}.` : "No data found here. If the data lives in a private repository, add its name and your token in Settings.");
    else if (state.loadError) text = "The roles couldn't load. Check your connection, then reload.";
    else if (state.loaded && !state.writable) text = "Read-only view. Add a GitHub token in Settings to save statuses and notes.";
    $("#noticeText").textContent = text;
    n.hidden = !text;
  }

  function render() {
    const a = document.activeElement;
    if (a && a.closest && a.closest("#roles") && /INPUT|TEXTAREA/.test(a.tagName)) { state.pendingRender = true; return; }
    state.pendingRender = false;
    const today = localISO();
    renderRunline();
    renderNotice();
    const all = merged();
    renderPipeline(all, today);
    $("#addBtn").hidden = !state.writable;
    const box = $("#roles");
    if (!state.loaded) {
      if (state.loadError) box.innerHTML = `<div class="empty"><h3>Roles aren't loaded</h3><p>${esc($("#noticeText").textContent)}</p></div>`;
      return;
    }
    if (!state.jobs.length) {
      box.innerHTML = `<div class="empty"><h3>No roles yet</h3><p>The twice-daily search adds verified openings to data/jobs.json. You can also add one you found with Add a role.</p></div>`;
      $("#count").textContent = "";
      return;
    }
    const rows = visibleRows();
    $("#count").textContent = `${rows.length} of ${all.length} roles · ${{ match: "best match first", newest: "newest first", company: "by company" }[state.f.sort]}`;
    box.innerHTML = rows.length ? rows.map((j) => roleHTML(j, today)).join("")
      : `<div class="empty"><h3>Nothing matches these filters</h3><p>Clear the search or switch filters back to All to see every role.</p></div>`;
  }

  // ---------------------------------------------------------------- toast
  let toastTimer = null;
  function toast(msg) {
    const t = $("#toast");
    t.textContent = msg; t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { t.hidden = true; }, 3400);
  }

  // ---------------------------------------------------------------- loading
  let loading = null;
  async function refresh() {
    if (loading) return loading;
    loading = (async () => {
      try {
        const [jobs, tracking, meta, profile] = await Promise.all([readJSON("jobs"), readJSON("tracking"), readJSON("meta"), readJSON("profile").catch(() => null)]);
        if (!jobs) { state.loadError = "notfound"; state.loaded = false; return; }
        state.jobs = Array.isArray(jobs.jobs) ? jobs.jobs : [];
        if (pendingWrites === 0) state.tracking = (tracking && tracking.tracking) || {};
        state.meta = meta || null;
        state.profile = profile || null;
        state.loaded = true; state.loadError = null;
        state.writable = canUseApi();
      } catch (e) {
        state.loadError = e.status === 0 ? "file" : (e.status === 401 || e.status === 403) ? "auth" : e.status === 404 ? "notfound" : "network";
        if (state.loadError === "auth") state.writable = false;
      } finally {
        loading = null;
        render();
      }
    })();
    return loading;
  }

  // ---------------------------------------------------------------- writes (serialized; always merged onto the latest file)
  let chain = Promise.resolve();
  const titleOf = (id) => { const j = state.jobs.find((x) => x.id === id); return j ? String(j.title).slice(0, 60).trim() : id; };
  function saveTracking(id, patch) {
    if (!state.writable) { toast("Add a GitHub token in Settings to save changes."); return Promise.resolve(false); }
    const stamp = new Date().toISOString();
    state.tracking[id] = Object.assign({}, state.tracking[id] || {}, patch, { updated_at: stamp });
    render();
    const job = async () => {
      for (let attempt = 0; attempt < 3; attempt++) {
        try {
          const latest = (await readJSON("tracking")) || { schema_version: 1, tracking: {} };
          latest.tracking = latest.tracking || {};
          latest.tracking[id] = Object.assign({}, latest.tracking[id] || {}, patch, { updated_at: stamp });
          const what = patch.status ? `status → ${patch.status}` : Object.keys(patch).join(", ");
          await writeJSON("tracking", latest, `Tracker: ${titleOf(id)} (${what})`);
          state.tracking = latest.tracking;
          render();
          return true;
        } catch (e) {
          if ((e.status === 409 || e.status === 422) && attempt < 2) continue;
          if (e.status === 401 || e.status === 403) toast("GitHub refused the save. The token needs Contents: Read and write on this repository.");
          else if (e.status === 404) toast("Couldn't find the repository. Check the owner and name in Settings.");
          else toast("That change didn't save. Check your connection and try again.");
          return false;
        }
      }
      return false;
    };
    pendingWrites++;
    const done = (ok) => { pendingWrites--; return ok; };
    chain = chain.then(job, job).then(done, () => done(false));
    return chain;
  }

  // ---------------------------------------------------------------- events
  document.addEventListener("click", (ev) => {
    const seg = ev.target.closest(".seg button");
    if (seg) {
      const key = seg.dataset.f, val = seg.dataset.v;
      state.f[key] = key === "match" ? Number(val) : val;
      seg.parentElement.querySelectorAll("button").forEach((b) => b.setAttribute("aria-pressed", String(b === seg)));
      render(); return;
    }
    const stage = ev.target.closest(".stage");
    if (stage) { const s = stage.dataset.stage; state.f.stage = state.f.stage === s ? null : s; render(); return; }
    if (ev.target.closest('[data-act="clearstage"]')) { state.f.stage = null; render(); }
  });
  document.addEventListener("toggle", (ev) => {
    const d = ev.target;
    if (d.matches && d.matches("details.more")) { if (d.open) state.open.add(d.dataset.id); else state.open.delete(d.dataset.id); }
  }, true);
  $("#q").addEventListener("input", (e) => { state.f.q = e.target.value; render(); });
  $("#showClosed").addEventListener("change", (e) => { state.f.showClosed = e.target.checked; render(); });
  $("#sort").addEventListener("change", (e) => { state.f.sort = e.target.value; render(); });

  $("#roles").addEventListener("change", (ev) => {
    const el = ev.target, act = el.dataset.act, art = el.closest("[data-id]");
    if (!act || !art) return;
    const id = art.dataset.id, cur = state.tracking[id] || {};
    if (act === "status") {
      el.dataset.s = el.value;
      const patch = { status: el.value };
      if (el.value === "Applied" && !cur.applied_on) patch.applied_on = localISO();
      saveTracking(id, patch).then((ok) => { if (ok) toast(`Status set to ${el.value}`); });
    } else if (act === "applied_on" || act === "contact" || act === "notes") {
      if ((cur[act] || "") === el.value) return;
      saveTracking(id, { [act]: el.value }).then((ok) => { if (ok) toast("Saved"); });
    }
  });
  $("#roles").addEventListener("focusout", () => { setTimeout(() => { if (state.pendingRender) render(); }, 0); });

  // settings
  function fillSettings() {
    $("#s-owner").value = cfg.owner; $("#s-repo").value = cfg.repo; $("#s-branch").value = cfg.branch;
    $("#s-token").value = cfg.token ? "••••••••" : "";
    $("#connLine").innerHTML = cfg.owner && cfg.repo
      ? `Data: <b>${esc(cfg.owner)}/${esc(cfg.repo)}</b> on <b>${esc(cfg.branch)}</b> · ${cfg.token ? "token saved on this device" : "no token (read-only)"}`
      : "Not connected yet.";
  }
  function openSettings() { fillSettings(); $("#settingsPanel").hidden = false; $("#s-owner").focus(); }
  $("#settingsBtn").addEventListener("click", () => { const p = $("#settingsPanel"); if (p.hidden) openSettings(); else p.hidden = true; });
  $("#noticeBtn").addEventListener("click", openSettings);
  $("#settingsCancel").addEventListener("click", () => { $("#settingsPanel").hidden = true; $("#settingsMsg").textContent = ""; });
  $("#forgetBtn").addEventListener("click", () => {
    cfg = Object.assign({}, cfg, { token: "" });
    storeSettings(cfg); fillSettings(); state.writable = false;
    $("#settingsMsg").textContent = "Token removed from this device.";
    refresh();
  });
  $("#settingsForm").addEventListener("submit", (ev) => {
    ev.preventDefault();
    const tokenInput = $("#s-token").value.trim();
    const next = {
      owner: $("#s-owner").value.trim(), repo: $("#s-repo").value.trim(), branch: $("#s-branch").value.trim() || "main",
      token: tokenInput === "••••••••" ? cfg.token : tokenInput,
    };
    if (!next.owner || !next.repo) { $("#settingsMsg").textContent = "Add the GitHub owner and the repository name."; return; }
    if (next.token && !/^(github_pat_|ghp_)[A-Za-z0-9_]+$/.test(next.token)) { $("#settingsMsg").textContent = "That doesn't look like a GitHub token (it starts with github_pat_ or ghp_)."; return; }
    cfg = next;
    if (!storeSettings(cfg)) { $("#settingsMsg").textContent = "This browser blocks saving settings (private mode?). They'll last until you close the tab."; }
    else $("#settingsMsg").textContent = "Saved.";
    state.loaded = false; state.loadError = null;
    fillSettings();
    refresh().then(() => { if (state.loaded) { $("#settingsPanel").hidden = true; toast(state.writable ? "Connected. Changes will save to GitHub." : "Connected (read-only)."); } });
  });

  // add a role
  const KEEP = new Set(["jid", "jl", "gh_jid", "id", "jobid", "job_id"]);
  function canonical(u) {
    const url = new URL(u.trim());
    const q = new URLSearchParams();
    for (const [k, v] of url.searchParams) if (KEEP.has(k.toLowerCase())) q.append(k, v);
    const path = url.pathname.replace(/\/+$/, "") || "/";
    const qs = q.toString();
    return `${url.protocol.replace(":", "").toLowerCase()}://${url.host.toLowerCase().replace(/^www\./, "")}${path}${qs ? "?" + qs : ""}`;
  }
  async function idFor(u) {
    const buf = await crypto.subtle.digest("SHA-1", new TextEncoder().encode(canonical(u)));
    return Array.from(new Uint8Array(buf)).map((b) => b.toString(16).padStart(2, "0")).join("").slice(0, 12);
  }
  $("#addBtn").addEventListener("click", () => { const p = $("#addPanel"); p.hidden = !p.hidden; if (!p.hidden) $("#f-title").focus(); });
  $("#addCancel").addEventListener("click", () => { $("#addPanel").hidden = true; $("#addMsg").textContent = ""; });
  $("#addForm").addEventListener("submit", async (ev) => {
    ev.preventDefault();
    const msg = $("#addMsg");
    const title = $("#f-title").value.trim(), company = $("#f-company").value.trim(), url = $("#f-url").value.trim();
    if (!title || !company || !/^https?:\/\/\S+\.\S+/i.test(url)) { msg.textContent = "Add a title, a company and the full posting link (starting with https://)."; return; }
    if (!state.writable) { msg.textContent = "Add a GitHub token in Settings first."; return; }
    msg.textContent = "Saving…";
    try {
      const id = await idFor(url);
      const today = localISO();
      const doc = {
        id, title, company, url, canonical_url: canonical(url),
        location: $("#f-location").value.trim() || "Not given", mode: $("#f-mode").value, type: $("#f-type").value,
        pay: $("#f-pay").value.trim() || "Not listed", posted: "Not shown", found_on: today, found_run: "", source: "Added by you",
        fit: null, tags: [], why: "", watch: "", auth: "Not stated", posting_status: "Open", last_checked: today, added_by: "manual", rank: 5000,
      };
      for (let attempt = 0; attempt < 3; attempt++) {
        const latest = (await readJSON("jobs")) || { schema_version: 1, jobs: [] };
        latest.jobs = Array.isArray(latest.jobs) ? latest.jobs : [];
        if (latest.jobs.some((j) => j.id === id)) { msg.textContent = "That posting is already in the tracker."; return; }
        latest.jobs.push(doc);
        latest.updated_at = new Date().toISOString();
        try { await writeJSON("jobs", latest, `Tracker: add ${title.slice(0, 60)} (${company.slice(0, 40)})`); state.jobs = latest.jobs; break; }
        catch (e) { if ((e.status === 409 || e.status === 422) && attempt < 2) continue; throw e; }
      }
      $("#addForm").reset(); msg.textContent = ""; $("#addPanel").hidden = true; render(); toast("Role added");
    } catch (e) {
      msg.textContent = e && (e.status === 401 || e.status === 403) ? "GitHub refused the save. Check the token in Settings." : "That role didn't save. Check the link and try again.";
    }
  });

  // excel
  function loadExcelJS() {
    if (window.ExcelJS) return Promise.resolve(window.ExcelJS);
    return new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = EXCELJS_URL; s.integrity = EXCELJS_SRI; s.crossOrigin = "anonymous";
      s.onload = () => (window.ExcelJS ? resolve(window.ExcelJS) : reject(new Error("ExcelJS missing")));
      s.onerror = () => reject(new Error("ExcelJS failed to load"));
      document.head.appendChild(s);
    });
  }
  $("#excelBtn").addEventListener("click", async () => {
    const btn = $("#excelBtn");
    if (!state.loaded) { toast("Load the roles first."); return; }
    btn.disabled = true; btn.textContent = "Preparing…";
    try {
      const ExcelJS = await loadExcelJS();
      const today = localISO();
      const wb = buildWorkbook(ExcelJS, sortRows(merged(), "match"), today, state.profile);
      const buf = await wb.xlsx.writeBuffer();
      const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = `Sumana_Job_Tracker_${today}.xlsx`;
      document.body.appendChild(a); a.click();
      setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
      toast("Excel downloaded");
    } catch (_) {
      toast("The Excel file couldn't be built. Check your connection and try again.");
    } finally {
      btn.disabled = false; btn.textContent = "Download Excel";
    }
  });

  // ---------------------------------------------------------------- boot
  render();
  refresh();
  setInterval(() => { $("#nextrun").textContent = nextRunText(); }, 60000);
  setInterval(() => { if (document.visibilityState === "visible") refresh(); }, 5 * 60000);
  document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible") refresh(); });
})();

/* Sumana's Job Tracker v2: Today / Opportunities / Pipeline / Contacts / Activity.
   Data lives as JSON in a GitHub repo (data/jobs.json, tracking.json, contacts.json, meta.json, profile.json).
   Reads go through the GitHub Contents API when a token is set (required for a private data repo);
   writes are small commits through the same API. */
"use strict";
(() => {
  const $ = (s, el = document) => el.querySelector(s);
  const $$ = (s, el = document) => Array.from(el.querySelectorAll(s));
  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const CT = "America/Chicago";
  const CLOSED = new Set(["Rejected", "Not a fit", "Closed"]);
  const TO_APPLY = new Set(["New", "Shortlisted"]);
  const WAITING = new Set(["Applied", "Recruiter screen", "Interviewing"]);
  const STALE_DAYS = 3;
  const FILES = { jobs: "data/jobs.json", tracking: "data/tracking.json", meta: "data/meta.json", profile: "data/profile.json", contacts: "data/contacts.json", reverify: "data/reverify.json" };
  const WORKFLOW = "reverify.yml";
  const LS_KEY = "sjt.settings.v1"; // shared with the classic view, so the token carries over
  const VIEW_KEY = "sjt.view.v2";
  const EXCELJS_URL = "https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js";
  const EXCELJS_SRI = "sha384-Pqp51FUN2/qzfxZxBCtF0stpc9ONI6MYZpVqmo8m20SoaQCzf+arZvACkLkirlPz";
  const VIEWS = { today: "Today", roles: "Opportunities", pipeline: "Pipeline", contacts: "Contacts", activity: "Activity" };
  const STAGES = [["New", "New"], ["Shortlisted", "Shortlisted"], ["Applied", "Applied"], ["Recruiter screen", "Screen"], ["Interviewing", "Interviewing"], ["Offer", "Offer"]];
  const BOARD = [["Shortlisted", "Ready to apply"], ["Applied", "Applied"], ["Recruiter screen", "Recruiter screen"], ["Interviewing", "Interviewing"], ["Offer", "Offer"]];
  const SRC_LABEL = { official: "Official", board: "Job board", vendor: "Vendor", recruiter: "Recruiter", referral: "Referral", chatgpt: "ChatGPT", manual: "Added by you" };

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
    return { owner: saved.owner || c.owner || d.owner, repo: saved.repo || c.dataRepo || d.repo, branch: saved.branch || c.branch || "main", token: saved.token || "" };
  }
  function storeSettings(s) { try { localStorage.setItem(LS_KEY, JSON.stringify(s)); return true; } catch (_) { return false; } }
  let cfg = loadSettings();
  const canUseApi = () => Boolean(cfg.token && cfg.owner && cfg.repo);
  const sameSiteRepo = () => { const d = detectRepo(); return d.owner && d.owner.toLowerCase() === (cfg.owner || "").toLowerCase() && d.repo.toLowerCase() === (cfg.repo || "").toLowerCase(); };

  // ---------------------------------------------------------------- GitHub data layer
  const API = "https://api.github.com";
  const shas = {};
  const b64enc = (str) => { const bytes = new TextEncoder().encode(str); let bin = ""; for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000)); return btoa(bin); };
  const b64dec = (b64) => { const bin = atob(String(b64).replace(/\s/g, "")); const bytes = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i); return new TextDecoder().decode(bytes); };
  const headers = (accept) => ({ Accept: accept || "application/vnd.github+json", "X-GitHub-Api-Version": "2022-11-28", Authorization: `Bearer ${cfg.token}` });
  const repoUrl = () => `${API}/repos/${encodeURIComponent(cfg.owner)}/${encodeURIComponent(cfg.repo)}`;
  function httpError(status, message) { const e = new Error(message || `HTTP ${status}`); e.status = status; return e; }

  async function readJSON(key) {
    const path = FILES[key];
    if (canUseApi()) {
      const url = `${repoUrl()}/contents/${path}?ref=${encodeURIComponent(cfg.branch)}`;
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
    if (cfg.owner && cfg.repo && detectRepo().owner && !sameSiteRepo()) throw httpError(404);
    const r = await fetch(`${path}?v=${Date.now()}`, { cache: "no-store" });
    if (r.status === 404) return null;
    if (!r.ok) throw httpError(r.status);
    return r.json();
  }
  async function writeJSON(key, obj, message) {
    const body = { message, content: b64enc(JSON.stringify(obj, null, 2) + "\n"), branch: cfg.branch };
    if (shas[key]) body.sha = shas[key];
    const r = await fetch(`${repoUrl()}/contents/${FILES[key]}`, { method: "PUT", headers: Object.assign(headers(), { "Content-Type": "application/json" }), body: JSON.stringify(body) });
    if (!r.ok) throw httpError(r.status);
    const out = await r.json();
    shas[key] = out && out.content ? out.content.sha : null;
  }
  // Read-modify-write against the latest file, retrying when someone else committed in between.
  async function commitChange(key, empty, mutate, message) {
    for (let attempt = 0; attempt < 3; attempt++) {
      const latest = (await readJSON(key)) || empty();
      const result = mutate(latest);
      if (result === false) return { latest, skipped: true };
      try { await writeJSON(key, latest, message); return { latest }; }
      catch (e) { if ((e.status === 409 || e.status === 422) && attempt < 2) continue; throw e; }
    }
    throw httpError(409);
  }
  async function listCommits() {
    if (!canUseApi()) return [];
    const r = await fetch(`${repoUrl()}/commits?sha=${encodeURIComponent(cfg.branch)}&per_page=40`, { headers: headers(), cache: "no-store" });
    if (!r.ok) throw httpError(r.status);
    const list = await r.json();
    return list.map((c) => ({ sha: c.sha, message: (c.commit && c.commit.message) || "", date: (c.commit && ((c.commit.committer && c.commit.committer.date) || (c.commit.author && c.commit.author.date))) || "" }));
  }

  // ---------------------------------------------------------------- state
  let pendingWrites = 0;
  const state = {
    jobs: [], tracking: {}, meta: null, profile: null, contacts: [], commits: null, reverify: null, recheckBusy: false,
    loaded: false, loadError: null, writable: false, view: "today",
    f: { q: "", where: "all", type: "all", match: 0, src: "", showClosed: false, sort: "match", stage: null },
    open: new Set(), pendingRender: false, editingContact: null,
  };
  try { const v = localStorage.getItem(VIEW_KEY); if (v && VIEWS[v]) state.view = v; } catch (_) { /* default view */ }
  if (VIEWS[location.hash.slice(1)]) state.view = location.hash.slice(1);

  // ---------------------------------------------------------------- dates
  const pad = (n) => String(n).padStart(2, "0");
  const localISO = (d = new Date()) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  const addDaysISO = (iso, n) => { const [y, m, d] = iso.split("-").map(Number); return localISO(new Date(y, m - 1, d + n)); };
  const shortDate = (iso) => { const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso || ""); if (!m) return iso || ""; return new Date(+m[1], +m[2] - 1, +m[3]).toLocaleDateString("en-US", { month: "short", day: "numeric" }); };
  const daysAgo = (iso, today) => { if (!iso) return null; const [a, b] = [iso, today].map((x) => { const [y, m, d] = x.split("-").map(Number); return Date.UTC(y, m - 1, d); }); return Math.round((b - a) / 864e5); };
  const fmtCT = (ts) => { const d = new Date(ts); return isNaN(d) ? "" : d.toLocaleString("en-US", { timeZone: CT, weekday: "short", month: "short", day: "numeric", hour: "numeric", minute: "2-digit" }); };
  const longToday = () => new Date().toLocaleDateString("en-US", { weekday: "long", month: "short", day: "numeric" });
  function nextRunText() {
    const p = Object.fromEntries(new Intl.DateTimeFormat("en-US", { timeZone: CT, hour: "numeric", minute: "numeric", hourCycle: "h23" }).formatToParts(new Date()).map((x) => [x.type, x.value]));
    const mins = (+p.hour) * 60 + (+p.minute);
    if (mins < 630) return "next search today 10:30 AM";
    if (mins < 1230) return "next search today 8:30 PM";
    return "next search tomorrow 10:30 AM";
  }
  function postedLabel(p) {
    const raw = String(p || "").trim();
    if (!raw || /^not (shown|stated|listed)$/i.test(raw)) return "";
    const s = raw.replace(/\d{4}-\d{2}-\d{2}/g, (iso) => shortDate(iso));
    if (/^\d+\+? days?$/i.test(s)) return `Posted ${s} ago`;
    return /^[~\d]|^[A-Z][a-z]{2} \d/.test(s) ? `Posted ${s}` : s;
  }

  // ---------------------------------------------------------------- derived data
  function merged() {
    return state.jobs.filter((j) => j && j.id).map((j) => {
      const t = state.tracking[j.id] || {};
      const override = t.posting_override && (!j.last_checked || (t.posting_override_on || "") >= j.last_checked) ? t.posting_override : "";
      const status = t.status || "New";
      return Object.assign({}, j, {
        // New and Shortlisted mean "not applied yet", so a leftover applied date is ignored
        status, applied_on: TO_APPLY.has(status) ? "" : t.applied_on || "", contact: t.contact || "", notes: t.notes || "",
        next_action: t.next_action || "", follow_up_on: t.follow_up_on || "", verified_on: t.verified_on || "",
        posting_status: override || j.posting_status,
      });
    });
  }
  const FOUND_BY = { "scheduled-search": "Claude", "initial-curation": "Claude", "chatgpt-search": "ChatGPT", "chatgpt-import": "ChatGPT", manual: "you" };
  const foundBy = (j) => FOUND_BY[j.added_by] || (/chatgpt/i.test(j.added_by || "") ? "ChatGPT" : /claude/i.test(j.added_by || "") ? "Claude" : "");
  const sourceType = (j) => j.source_type || (j.added_by === "manual" ? "manual" : /\(vendor\)/i.test(j.company || "") ? "vendor" : "board");
  const verifiedOn = (j) => [j.last_checked, j.verified_on].filter(Boolean).sort().pop() || "";
  const isPassed = (j) => CLOSED.has(j.status) || j.posting_status === "Closed";
  const isStale = (j, today) => TO_APPLY.has(j.status) && !isPassed(j) && /^https?:/i.test(j.url || "") && (!verifiedOn(j) || addDaysISO(verifiedOn(j), STALE_DAYS) < today);
  const followDate = (j) => j.follow_up_on || (j.applied_on ? addDaysISO(j.applied_on, 7) : "");
  const followDue = (j, today) => WAITING.has(j.status) && !!followDate(j) && followDate(j) <= today;
  const contactDue = (c, today) => c.status !== "Done" && !!c.next_follow_up && c.next_follow_up <= today;
  const isChicago = (j) => j.mode === "Hybrid" || j.mode === "On-site" || /chicago|naperville|illinois|,\s*IL\b|schaumburg|rolling meadows|batavia|lemont|itasca/i.test(j.location || "");
  const isRemote = (j) => j.mode === "Remote" || j.mode === "Confirm";
  const suggestedAction = (j) => j.next_action || (
    j.status === "New" ? ((j.fit || 0) >= 5 ? "Tailor resume, then apply" : "Review and shortlist") :
    j.status === "Shortlisted" ? "Tailor resume, then apply" :
    j.status === "Applied" ? "Follow up 7 days after applying" :
    j.status === "Recruiter screen" ? "Follow up with recruiter" :
    j.status === "Interviewing" ? "Prep for interview" :
    j.status === "Offer" ? "Review offer" : "");
  const priority = (j) => (j.fit || 0) * 10 + (j.status === "Shortlisted" ? 4 : 0) + (["recruiter", "referral"].includes(sourceType(j)) ? 3 : 0) + (j.mode === "Remote" ? 1 : 0);
  function sortRows(rows, how) {
    const byMatch = (a, b) => (b.fit || 0) - (a.fit || 0) || String(b.found_on || "").localeCompare(String(a.found_on || "")) || (a.rank ?? 9999) - (b.rank ?? 9999);
    if (how === "newest") return rows.sort((a, b) => String(b.found_on || "").localeCompare(String(a.found_on || "")) || String(b.found_run || "").localeCompare(String(a.found_run || "")) || byMatch(a, b));
    if (how === "company") return rows.sort((a, b) => String(a.company || "").localeCompare(String(b.company || "")) || byMatch(a, b));
    if (how === "priority") return rows.sort((a, b) => priority(b) - priority(a) || byMatch(a, b));
    return rows.sort(byMatch);
  }

  // ---------------------------------------------------------------- small render helpers
  const pips = (n) => Array.from({ length: 5 }, (_, i) => `<span class="pip${i < (n || 0) ? " on" : ""}"></span>`).join("");
  const srcBadge = (j) => { const t = sourceType(j); return `<span class="src src-${esc(t)}">${esc(SRC_LABEL[t] || t)}</span>`; };
  const fitBadge = (j) => `<span class="fitb">${j.fit ? `${esc(j.fit)}/5` : "—"}</span>`;
  const openLink = (j, label = "Open posting ↗") => /^https?:\/\//i.test(j.url || "") ? `<a class="qbtn" href="${esc(j.url)}" target="_blank" rel="noopener noreferrer">${esc(label)}</a>` : "";
  const dis = () => (state.writable ? "" : " disabled");
  const autoNote = (j) => { const r = state.reverify && state.reverify.results && state.reverify.results[j.id]; return r && r.result === "unknown" && verifiedOn(j) <= String(r.checked_at || "").slice(0, 10) ? r.note : ""; };

  // ---------------------------------------------------------------- TODAY
  function renderToday(all, today) {
    const active = all.filter((j) => !isPassed(j));
    const stale = sortRows(active.filter((j) => isStale(j, today)), "match");
    const applyNext = sortRows(active.filter((j) => TO_APPLY.has(j.status) && !isStale(j, today)), "priority").slice(0, 5);
    const follow = sortRows(active.filter((j) => followDue(j, today)), "priority");
    const contactsDue = state.contacts.filter((c) => contactDue(c, today));
    const fit5 = active.filter((j) => j.fit === 5 && TO_APPLY.has(j.status)).length;
    const dueCount = follow.length + contactsDue.length;
    $("#kpis").innerHTML = [
      ["Active roles", active.length, "", "Open and not passed"],
      ["Match 5 to apply", fit5, "", "Not applied yet"],
      ["Follow-ups due", dueCount, dueCount ? "warn" : "", "Roles and contacts"],
      ["Re-check first", stale.length, stale.length ? "check" : "", `Verified ${STALE_DAYS}+ days ago`],
    ].map(([l, n, cls, s]) => `<div class="kpi ${cls}"><span class="l">${esc(l)}</span><span class="n">${n}</span><span class="s">${esc(s)}</span></div>`).join("");

    const writable = state.writable;
    $("#laneApply").innerHTML = applyNext.length ? applyNext.map((j) => `
      <div class="item" data-id="${esc(j.id)}">
        <div class="t">${esc(j.title)}</div>
        <div class="s">${esc(j.company)} · ${esc(j.location)}</div>
        <div class="badges">${fitBadge(j)}${srcBadge(j)}<span class="chip">${esc(j.type || "")}</span></div>
        <div class="na">${esc(suggestedAction(j))}</div>
        <div class="quick">${openLink(j)}${writable ? `<button type="button" class="qbtn primary" data-act="mark-applied">Mark applied</button>${j.status === "New" ? '<button type="button" class="qbtn" data-act="set-status" data-v="Shortlisted">Shortlist</button>' : ""}<button type="button" class="qbtn" data-act="set-status" data-v="Not a fit">Not a fit</button>` : ""}</div>
      </div>`).join("") : `<p class="empty-s">Nothing waiting. New roles land here after each search.</p>`;

    const followItems = follow.map((j) => `
      <div class="item" data-id="${esc(j.id)}">
        <div class="t">${esc(j.title)}</div>
        <div class="s">${esc(j.company)} · ${esc(j.status)} · due ${esc(shortDate(followDate(j)))}</div>
        ${j.contact ? `<div class="s">Contact: ${esc(j.contact)}</div>` : ""}
        <div class="quick">${writable ? '<button type="button" class="qbtn primary" data-act="followed-up">Followed up</button>' : ""}<button type="button" class="qbtn" data-act="open-role">Notes</button></div>
      </div>`).join("");
    const contactItems = contactsDue.map((c) => `
      <div class="item" data-cid="${esc(c.id)}">
        <div class="t">${esc(c.name)} <span class="src src-recruiter">${esc(c.relationship || "Contact")}</span></div>
        <div class="s">${esc(c.company || "")}${c.related_role ? ` · ${esc(c.related_role)}` : ""}</div>
        <div class="s">Follow-up due ${esc(shortDate(c.next_follow_up))}${c.reach ? ` · ${esc(c.reach)}` : ""}</div>
        <div class="quick">${writable ? '<button type="button" class="qbtn primary" data-act="contact-followed">Followed up</button>' : ""}<button type="button" class="qbtn" data-act="contact-edit">Open</button></div>
      </div>`).join("");
    $("#laneFollow").innerHTML = (contactItems + followItems) || `<p class="empty-s">No follow-ups due. Applied roles show up here 7 days after applying.</p>`;

    $("#laneCheck").innerHTML = stale.length ? stale.slice(0, 6).map((j) => `
      <div class="item" data-id="${esc(j.id)}">
        <div class="t">${esc(j.title)}</div>
        <div class="s">${esc(j.company)} · last verified ${verifiedOn(j) ? `${daysAgo(verifiedOn(j), today)} days ago` : "never"}</div>
        ${autoNote(j) ? `<div class="s autonote">Auto-check couldn't confirm: ${esc(autoNote(j))}. Open it to check.</div>` : ""}
        <div class="quick">${openLink(j)}${writable ? `<button type="button" class="qbtn primary" data-act="still-open">Still open</button><button type="button" class="qbtn" data-act="set-status" data-v="Closed">Closed</button>${autoNote(j) ? "" : '<button type="button" class="qbtn" data-act="auto-check">Auto-check</button>'}` : ""}</div>
      </div>`).join("") + (stale.length > 6 ? `<p class="empty-s">${stale.length - 6} more in Opportunities.</p>` : "")
      : `<p class="empty-s">Everything was verified in the last ${STALE_DAYS} days.</p>`;

    const top = sortRows(active.filter((j) => TO_APPLY.has(j.status) || WAITING.has(j.status)), "priority").slice(0, 8);
    $("#topBody").innerHTML = top.length ? top.map((j) => `<tr data-id="${esc(j.id)}">
        <td><span class="pips" aria-label="Match ${esc(j.fit || "not scored")} of 5">${pips(j.fit)}</span></td>
        <td class="wrapcell"><button type="button" class="linkrow" data-act="open-role">${esc(j.title)}</button></td>
        <td class="wrapcell">${esc(j.company)}</td>
        <td>${esc(j.mode === "Confirm" ? "Confirm" : j.mode)}</td>
        <td class="wrapcell">${esc(j.type || "")}${j.pay && j.pay !== "Not listed" ? ` · ${esc(j.pay)}` : ""}</td>
        <td>${srcBadge(j)}</td>
        <td>${esc(j.status)}</td>
        <td class="wrapcell">${esc(suggestedAction(j))}</td></tr>`).join("")
      : `<tr><td colspan="8">No active roles yet.</td></tr>`;
    return dueCount;
  }

  // ---------------------------------------------------------------- OPPORTUNITIES
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
      if (f.src && sourceType(j) !== f.src) return false;
      if (q && ![j.title, j.company, j.location, j.type, j.why, j.req_id, (j.tags || []).join(" ")].join(" ").toLowerCase().includes(q)) return false;
      return true;
    });
    return sortRows(rows, f.sort);
  }
  function roleHTML(j, today) {
    const id = esc(j.id);
    const isNew = j.found_on && j.found_on >= addDaysISO(today, -1) && j.status === "New";
    const opts = STATUSES.map((s) => `<option${s === j.status ? " selected" : ""}>${esc(s)}</option>`).join("");
    const pl = postedLabel(j.posted);
    const chips = [
      j.mode ? `<span class="chip mode-${esc(String(j.mode).replace(/\s+/g, "-"))}">${esc(j.mode === "Confirm" ? "Remote? confirm" : j.mode)}</span>` : "",
      j.type ? `<span class="chip">${esc(j.type)}</span>` : "",
      j.pay && j.pay !== "Not listed" ? `<span class="chip">${esc(j.pay)}</span>` : "",
      j.req_id ? `<span class="chip quiet">Req ${esc(j.req_id)}</span>` : "",
      pl ? `<span class="chip quiet">${esc(pl)}</span>` : "",
      j.found_on ? `<span class="chip quiet">Found ${esc(shortDate(j.found_on))}</span>` : "",
    ].join("");
    const tags = (Array.isArray(j.tags) ? j.tags : []).map((t) => `<span class="tag">${esc(t)}</span>`).join("");
    const fd = followDate(j);
    const track = j.applied_on ? `<div class="trackline">Applied ${esc(shortDate(j.applied_on))}${fd ? ` · follow up by ${esc(shortDate(fd))}` : ""}</div>` : "";
    const ver = verifiedOn(j);
    return `<article class="role${isPassed(j) ? " is-dim" : ""}" data-id="${id}">
      <div class="match" aria-label="Match ${esc(j.fit || "not scored")} of 5"><span class="pips" aria-hidden="true">${pips(j.fit)}</span><span class="mlabel">Match <b>${esc(j.fit || "–")}</b></span></div>
      <div class="rbody">
        <div class="titleline"><h3>${esc(j.title)}</h3>${isNew ? '<span class="badge badge-new">New</span>' : ""}${j.posting_status === "Closed" ? '<span class="badge badge-closed">Posting closed</span>' : ""}${followDue(j, today) ? '<span class="badge badge-due">Follow up</span>' : ""}${isStale(j, today) ? '<span class="badge badge-stale">Re-check</span>' : ""}</div>
        <div class="org"><b>${esc(j.company)}</b> · ${esc(j.location)} ${srcBadge(j)}</div>
        <div class="chips">${chips}</div>
        ${tags ? `<div class="tags" aria-label="Skill match">${tags}</div>` : ""}
        <div class="na">Next: ${esc(suggestedAction(j) || "—")}</div>
        ${autoNote(j) ? `<div class="autonote">Auto-check couldn't confirm: ${esc(autoNote(j))}. Open the posting to check.</div>` : ""}
      </div>
      <div class="side-ctl">
        <label class="sr-only" for="st-${id}">Status for ${esc(j.title)}</label>
        <select class="status" id="st-${id}" data-s="${esc(j.status)}" data-act="status"${dis()}>${opts}</select>
        ${/^https?:\/\//i.test(j.url || "") ? `<a class="open" href="${esc(j.url)}" target="_blank" rel="noopener noreferrer">Open posting ↗</a>` : '<span class="trackline">No posting link (lead)</span>'}
        ${j.posting_status === "Closed" && state.writable ? '<button type="button" class="qbtn" data-act="reopen">Still open? Reopen</button>' : ""}
        ${track}
      </div>
      <details class="more" data-id="${id}"${state.open.has(j.id) ? " open" : ""}>
        <summary>Why it fits, next action and notes</summary>
        <div class="more-grid">
          <div><h4>Why it fits</h4><p>${esc(j.why || "—")}</p></div>
          <div><h4>Watch-outs</h4><p>${esc(j.watch || "—")}</p></div>
          <div><h4>Work authorization</h4><p>${esc(j.auth || "Not stated")}</p></div>
          <div><h4>Source</h4><p>${esc(j.source || "—")}${foundBy(j) ? ` · found by ${esc(foundBy(j))}` : ""}${ver ? ` · verified ${esc(shortDate(ver))}` : ""}</p></div>
          <div class="track-grid">
            <div class="field full"><label for="na-${id}">Next action</label><input id="na-${id}" data-act="next_action" list="nextActions" value="${esc(j.next_action)}" placeholder="${esc(suggestedAction(j))}"${dis()}></div>
            <div class="field"><label for="ap-${id}">Applied on</label><input type="date" id="ap-${id}" data-act="applied_on" value="${esc(j.applied_on)}"${dis()}></div>
            <div class="field"><label for="fu-${id}">Follow up by</label><input type="date" id="fu-${id}" data-act="follow_up_on" value="${esc(j.follow_up_on)}" placeholder="${esc(fd)}"${dis()}></div>
            <div class="field full"><label for="ct-${id}">Recruiter / contact</label><input id="ct-${id}" data-act="contact" value="${esc(j.contact)}" placeholder="Name, vendor, phone or email"${dis()}></div>
            <div class="field full"><label for="nt-${id}">Notes</label><textarea id="nt-${id}" data-act="notes" placeholder="Resume version sent, rate discussed, interview dates…"${dis()}>${esc(j.notes)}</textarea></div>
          </div>
        </div>
      </details>
    </article>`;
  }
  function renderRoles(all, today) {
    const counts = {};
    all.forEach((j) => { counts[j.status] = (counts[j.status] || 0) + 1; });
    $("#pipeline").innerHTML = STAGES.map(([s, label]) => `<button type="button" class="stage" data-stage="${esc(s)}" aria-pressed="${state.f.stage === s}"><span class="n">${counts[s] || 0}</span><span class="l">${esc(label)}</span></button>`).join("");
    const applied = all.filter((j) => j.applied_on).length;
    const responded = all.filter((j) => j.applied_on && ["Recruiter screen", "Interviewing", "Offer"].includes(j.status)).length;
    const bits = [];
    if (applied) bits.push(`<span>${applied} applied · ${Math.round((responded / applied) * 100)}% response rate</span>`);
    bits.push(`<span>${all.filter(isPassed).length} closed or passed</span>`);
    if (state.f.stage) bits.push(`<button type="button" class="linkish" data-act="clearstage">Show all statuses</button>`);
    $("#pipenotes").innerHTML = bits.join("");
    const box = $("#roles");
    if (!all.length) { box.innerHTML = `<div class="empty"><h3>No roles yet</h3><p>The twice-daily search adds verified openings here. You can also use Add a role.</p></div>`; $("#count").textContent = ""; return; }
    const rows = visibleRows();
    $("#count").textContent = `${rows.length} of ${all.length} roles`;
    box.innerHTML = rows.length ? rows.map((j) => roleHTML(j, today)).join("") : `<div class="empty"><h3>Nothing matches these filters</h3><p>Clear the search or switch filters back to All.</p></div>`;
  }

  // ---------------------------------------------------------------- PIPELINE
  function renderBoard(all) {
    $("#board").innerHTML = BOARD.map(([status, label]) => {
      const cards = sortRows(all.filter((j) => j.status === status && j.posting_status !== "Closed"), "priority");
      return `<section class="col" data-status="${esc(status)}" aria-label="${esc(label)}">
        <h3><span>${esc(label)}</span><span>${cards.length}</span></h3>
        ${cards.map((j) => `<article class="kcard" draggable="${state.writable}" data-id="${esc(j.id)}">
          <div class="kt">${esc(j.title)}</div>
          <div class="ks">${esc(j.company)}</div>
          <div class="badges"><span class="pips" aria-label="Match ${esc(j.fit || "not scored")} of 5">${pips(j.fit)}</span>${srcBadge(j)}</div>
          <div class="na">${esc(suggestedAction(j))}</div>
          <label class="sr-only" for="kb-${esc(j.id)}">Status</label>
          <select class="status" id="kb-${esc(j.id)}" data-s="${esc(j.status)}" data-act="status"${dis()}>${STATUSES.map((s) => `<option${s === j.status ? " selected" : ""}>${esc(s)}</option>`).join("")}</select>
        </article>`).join("") || '<p class="empty-s">Nothing here yet.</p>'}
      </section>`;
    }).join("");
    const newCount = all.filter((j) => j.status === "New" && !isPassed(j)).length;
    $("#boardNote").innerHTML = `<span>${newCount} new role${newCount === 1 ? "" : "s"} not yet shortlisted</span><span>${all.filter(isPassed).length} closed or passed</span>`;
  }

  // ---------------------------------------------------------------- CONTACTS
  function renderContacts(today) {
    const order = { Active: 0, Waiting: 1, Done: 2 };
    const list = state.contacts.slice().sort((a, b) => (contactDue(b, today) - contactDue(a, today)) || ((order[a.status] ?? 3) - (order[b.status] ?? 3)) || String(a.name || "").localeCompare(String(b.name || "")));
    $("#contacts").innerHTML = list.length ? list.map((c) => `
      <article class="ccard${c.status === "Done" ? " is-dim" : ""}" data-cid="${esc(c.id)}">
        <div class="cmain">
          <div class="titleline"><h3>${esc(c.name)}</h3><span class="src src-recruiter">${esc(c.relationship || "Contact")}</span>${contactDue(c, today) ? '<span class="badge badge-due">Follow up</span>' : ""}<span class="chip quiet">${esc(c.status || "Active")}</span></div>
          <div class="org">${c.company ? `<b>${esc(c.company)}</b>` : ""}${c.related_role ? ` · ${esc(c.related_role)}` : ""}</div>
          <div class="cmeta">${c.reach ? `<span>${esc(c.reach)}</span>` : ""}${c.last_contact ? `<span>Last contact ${esc(shortDate(c.last_contact))}</span>` : ""}${c.next_follow_up ? `<span>Next follow-up ${esc(shortDate(c.next_follow_up))}</span>` : ""}</div>
          ${c.notes ? `<p class="cnotes">${esc(c.notes)}</p>` : ""}
        </div>
        <div class="quick">${state.writable ? `<button type="button" class="qbtn primary" data-act="contact-followed">Followed up</button><button type="button" class="qbtn" data-act="contact-edit">Edit</button>${c.status !== "Done" ? '<button type="button" class="qbtn" data-act="contact-done">Mark done</button>' : ""}` : ""}</div>
      </article>`).join("")
      : `<div class="empty"><h3>No contacts yet</h3><p>Add recruiters, referrals and hiring managers here. Their follow-up dates show up on Today.</p></div>`;
    $("#roleTitles").innerHTML = state.jobs.map((j) => `<option value="${esc(`${j.title} — ${j.company}`)}"></option>`).join("");
  }

  // ---------------------------------------------------------------- ACTIVITY
  function renderActivity() {
    const runs = state.meta && Array.isArray(state.meta.runs) ? state.meta.runs.slice(0, 14) : [];
    $("#logBody").innerHTML = runs.length
      ? runs.map((r) => `<tr><td>${esc(r.label || fmtCT(r.at))}</td><td>${esc(r.new ?? "–")}</td><td>${esc(r.checked ?? "–")}</td><td>${esc(r.closed ?? "–")}</td></tr>`).join("")
      : `<tr><td colspan="4">No runs recorded yet.</td></tr>`;
    const tl = $("#timeline");
    if (!canUseApi()) { tl.innerHTML = `<p class="empty-s">Add your GitHub token in Settings to see the history.</p>`; return; }
    if (state.commits === null) { tl.innerHTML = `<p class="empty-s">Loading history…</p>`; loadCommits(); return; }
    tl.innerHTML = state.commits.length ? state.commits.map((c) => {
      const first = c.message.split("\n")[0];
      const kind = /^Job search/i.test(first) ? "search" : /^Tracker:/i.test(first) ? "update" : "other";
      const text = first.replace(/^Tracker:\s*/i, "");
      return `<div class="ev ev-${kind}"><b>${esc(text)}</b><span class="when">${esc(fmtCT(c.date))}${kind === "search" ? " · twice-daily search" : kind === "update" ? " · tracker change" : ""}</span></div>`;
    }).join("") : `<p class="empty-s">No history yet.</p>`;
  }
  let commitsLoading = false;
  async function loadCommits() {
    if (commitsLoading) return;
    commitsLoading = true;
    try { state.commits = await listCommits(); }
    catch (_) { state.commits = []; toast("Couldn't load the history from GitHub."); }
    finally { commitsLoading = false; if (state.view === "activity") renderActivity(); }
  }

  // ---------------------------------------------------------------- shell
  function renderNotice() {
    let text = "";
    if (state.loadError === "file") text = "Open this tracker through GitHub Pages or a local web server; browsers block data files opened straight from disk.";
    else if (state.loadError === "auth") text = "GitHub didn't accept the saved token. Check it in Settings.";
    else if (state.loadError === "notfound") text = canUseApi() ? "No data files found in that repository. Check the owner, repository and branch in Settings." : (cfg.owner && cfg.repo ? `Add your GitHub token in Settings to load the roles from ${cfg.owner}/${cfg.repo}.` : "No data found here. Add the data repository and your token in Settings.");
    else if (state.loadError) text = "The roles couldn't load. Check your connection, then reload.";
    else if (state.loaded && !state.writable) text = "Read-only view. Add a GitHub token in Settings to save changes.";
    $("#noticeText").textContent = text;
    $("#notice").hidden = !text;
  }
  function setView(v) {
    if (!VIEWS[v]) return;
    state.view = v;
    try { localStorage.setItem(VIEW_KEY, v); } catch (_) { /* fine */ }
    if (location.hash !== `#${v}`) history.replaceState(null, "", `#${v}`);
    render();
    window.scrollTo({ top: 0 });
  }
  function render() {
    const a = document.activeElement;
    if (a && a.closest && a.closest(".view") && /INPUT|TEXTAREA/.test(a.tagName) && !a.closest(".filters")) { state.pendingRender = true; return; }
    state.pendingRender = false;
    const today = localISO();
    $$(".navbtn").forEach((b) => { if (b.dataset.view === state.view) b.setAttribute("aria-current", "page"); else b.removeAttribute("aria-current"); });
    Object.keys(VIEWS).forEach((v) => { $(`#view-${v}`).hidden = v !== state.view; });
    $("#viewTitle").textContent = VIEWS[state.view];
    $("#viewSub").textContent = `${longToday()} · ${nextRunText()}`;
    const m = state.meta;
    $("#lastrun").innerHTML = m && m.last_run_at ? `Last run <b>${esc(fmtCT(m.last_run_at))}</b>${typeof m.last_new_count === "number" ? ` · ${m.last_new_count} new` : ""}` : "";
    $("#addBtn").hidden = !state.writable;
    $("#recheckBtn").hidden = !state.writable;
    renderRecheck();
    renderNotice();
    const all = merged();
    // nav counters
    const due = all.filter((j) => !isPassed(j) && followDue(j, today)).length + state.contacts.filter((c) => contactDue(c, today)).length;
    const setCount = (id, n) => { const el = $(id); el.textContent = n; el.hidden = !n; };
    setCount("#nc-today", due);
    setCount("#nc-roles", all.filter((j) => !isPassed(j)).length);
    setCount("#nc-contacts", state.contacts.filter((c) => contactDue(c, today)).length);
    if (!state.loaded) {
      if (state.loadError) $("#roles").innerHTML = `<div class="empty"><h3>Roles aren't loaded</h3><p>${esc($("#noticeText").textContent)}</p></div>`;
      if (state.view === "today") { $("#kpis").innerHTML = ""; ["#laneApply", "#laneFollow", "#laneCheck"].forEach((s) => { $(s).innerHTML = `<p class="empty-s">${state.loadError ? "Connect in Settings to load." : "Loading…"}</p>`; }); $("#topBody").innerHTML = ""; }
      return;
    }
    if (state.view === "today") renderToday(all, today);
    else if (state.view === "roles") renderRoles(all, today);
    else if (state.view === "pipeline") renderBoard(all);
    else if (state.view === "contacts") renderContacts(today);
    else if (state.view === "activity") renderActivity();
  }

  // ---------------------------------------------------------------- toast
  let toastTimer = null;
  function toast(msg) { const t = $("#toast"); t.textContent = msg; t.hidden = false; clearTimeout(toastTimer); toastTimer = setTimeout(() => { t.hidden = true; }, 3400); }
  function saveError(e) {
    if (e && (e.status === 401 || e.status === 403)) toast("GitHub refused the save. The token needs Contents: Read and write on the data repository.");
    else if (e && e.status === 404) toast("Couldn't find the data repository. Check Settings.");
    else toast("That change didn't save. Check your connection and try again.");
  }

  // ---------------------------------------------------------------- loading
  let loading = null;
  async function refresh() {
    if (loading) return loading;
    loading = (async () => {
      try {
        const [jobs, tracking, meta, profile, contacts, reverify] = await Promise.all([
          readJSON("jobs"), readJSON("tracking"), readJSON("meta"),
          readJSON("profile").catch(() => null), readJSON("contacts").catch(() => null), readJSON("reverify").catch(() => null),
        ]);
        if (!jobs) { state.loadError = "notfound"; state.loaded = false; return; }
        state.jobs = Array.isArray(jobs.jobs) ? jobs.jobs : [];
        if (pendingWrites === 0) {
          state.tracking = (tracking && tracking.tracking) || {};
          state.contacts = (contacts && Array.isArray(contacts.contacts)) ? contacts.contacts : [];
        }
        state.meta = meta || null;
        state.profile = profile || null;
        state.reverify = reverify || null;
        state.loaded = true; state.loadError = null;
        state.writable = canUseApi();
        if (state.view === "activity") state.commits = null;
      } catch (e) {
        state.loadError = e.status === 0 ? "file" : (e.status === 401 || e.status === 403) ? "auth" : e.status === 404 ? "notfound" : "network";
        if (state.loadError === "auth") state.writable = false;
      } finally { loading = null; render(); }
    })();
    return loading;
  }

  // ---------------------------------------------------------------- writes
  let chain = Promise.resolve();
  const titleOf = (id) => { const j = state.jobs.find((x) => x.id === id); return j ? String(j.title).slice(0, 60).trim() : id; };
  function queue(fn) {
    pendingWrites++;
    const run = () => fn().then((ok) => { pendingWrites--; return ok; }, (e) => { pendingWrites--; saveError(e); return false; });
    chain = chain.then(run, run);
    return chain;
  }
  // Status change plus the applied date that goes with it. New/Shortlisted mean "not applied yet":
  // moving into Applied from there stamps today, and any leftover applied date is cleared.
  function statusPatch(cur, status) {
    const patch = { status };
    const fromToApply = TO_APPLY.has(cur.status || "New");
    if (status === "Applied" && (fromToApply || !cur.applied_on)) patch.applied_on = localISO();
    else if (cur.applied_on && (fromToApply || TO_APPLY.has(status))) patch.applied_on = "";
    return patch;
  }
  function saveTracking(id, patch, label) {
    if (!state.writable) { toast("Add a GitHub token in Settings to save changes."); return Promise.resolve(false); }
    const stamp = new Date().toISOString();
    state.tracking[id] = Object.assign({}, state.tracking[id] || {}, patch, { updated_at: stamp, updated_by: "site" });
    render();
    const what = label || (patch.status ? `status → ${patch.status}` : Object.keys(patch).join(", "));
    return queue(async () => {
      const { latest } = await commitChange("tracking", () => ({ schema_version: 1, tracking: {} }), (doc) => {
        doc.tracking = doc.tracking || {};
        doc.tracking[id] = Object.assign({}, doc.tracking[id] || {}, patch, { updated_at: stamp, updated_by: "site" });
      }, `Tracker: ${titleOf(id)} (${what})`);
      state.tracking = latest.tracking; render();
      return true;
    });
  }
  function saveContact(contact, label) {
    if (!state.writable) { toast("Add a GitHub token in Settings to save changes."); return Promise.resolve(false); }
    contact.updated_at = new Date().toISOString();
    const idx = state.contacts.findIndex((c) => c.id === contact.id);
    if (idx >= 0) state.contacts[idx] = contact; else state.contacts.push(contact);
    render();
    return queue(async () => {
      const { latest } = await commitChange("contacts", () => ({ schema_version: 1, contacts: [] }), (doc) => {
        doc.contacts = Array.isArray(doc.contacts) ? doc.contacts : [];
        const i = doc.contacts.findIndex((c) => c.id === contact.id);
        if (i >= 0) doc.contacts[i] = Object.assign({}, doc.contacts[i], contact); else doc.contacts.push(contact);
      }, `Tracker: contact ${String(contact.name || "").slice(0, 40)} (${label})`);
      state.contacts = latest.contacts; render();
      return true;
    });
  }

  // ---------------------------------------------------------------- events
  document.addEventListener("click", (ev) => {
    const nav = ev.target.closest(".navbtn");
    if (nav) { setView(nav.dataset.view); return; }
    const go = ev.target.closest("[data-go]");
    if (go) { setView(go.dataset.go); return; }
    const seg = ev.target.closest(".seg button");
    if (seg) {
      const key = seg.dataset.f, val = seg.dataset.v;
      state.f[key] = key === "match" ? Number(val) : val;
      seg.parentElement.querySelectorAll("button").forEach((b) => b.setAttribute("aria-pressed", String(b === seg)));
      render(); return;
    }
    const stage = ev.target.closest(".stage");
    if (stage) { const s = stage.dataset.stage; state.f.stage = state.f.stage === s ? null : s; render(); return; }
    const act = ev.target.closest("[data-act]");
    if (!act || act.tagName === "SELECT" || act.tagName === "INPUT" || act.tagName === "TEXTAREA") return;
    const a = act.dataset.act;
    const holder = act.closest("[data-id]");
    const id = holder && holder.dataset.id;
    const cHolder = act.closest("[data-cid]");
    const cid = cHolder && cHolder.dataset.cid;
    const today = localISO();
    if (a === "clearstage") { state.f.stage = null; render(); }
    else if (a === "mark-applied" && id) { saveTracking(id, statusPatch(state.tracking[id] || {}, "Applied")).then((ok) => ok && toast("Marked applied. Follow-up in 7 days.")); }
    else if (a === "set-status" && id) { saveTracking(id, statusPatch(state.tracking[id] || {}, act.dataset.v)).then((ok) => ok && toast(`Status set to ${act.dataset.v}`)); }
    else if (a === "auto-check" && id) { startRecheck("stale", [id]); }
    else if (a === "reopen" && id) { saveTracking(id, { posting_override: "Open", posting_override_on: today, verified_on: today }, "reopened").then((ok) => ok && toast("Marked as still open")); }
    else if (a === "still-open" && id) { saveTracking(id, { verified_on: today }, "still open").then((ok) => ok && toast("Marked as verified today")); }
    else if (a === "followed-up" && id) { saveTracking(id, { follow_up_on: addDaysISO(today, 7) }, "followed up").then((ok) => ok && toast("Next follow-up in 7 days")); }
    else if (a === "open-role" && id) { state.open.add(id); state.f.stage = null; state.f.q = ""; $("#q").value = ""; setView("roles"); setTimeout(() => { const el = document.querySelector(`article.role[data-id="${CSS.escape(id)}"]`); if (el) el.scrollIntoView({ block: "center" }); }, 50); }
    else if (a === "contact-followed" && cid) { const c = state.contacts.find((x) => x.id === cid); if (c) saveContact(Object.assign({}, c, { last_contact: today, next_follow_up: addDaysISO(today, 7) }), "followed up").then((ok) => ok && toast("Logged. Next follow-up in 7 days.")); }
    else if (a === "contact-done" && cid) { const c = state.contacts.find((x) => x.id === cid); if (c) saveContact(Object.assign({}, c, { status: "Done" }), "done").then((ok) => ok && toast("Contact marked done")); }
    else if (a === "contact-edit" && cid) { if (state.view !== "contacts") setView("contacts"); openContactForm(state.contacts.find((x) => x.id === cid)); }
  });
  document.addEventListener("toggle", (ev) => {
    const d = ev.target;
    if (d.matches && d.matches("details.more")) { if (d.open) state.open.add(d.dataset.id); else state.open.delete(d.dataset.id); }
  }, true);
  document.addEventListener("change", (ev) => {
    const el = ev.target, act = el.dataset && el.dataset.act;
    const holder = el.closest && el.closest("[data-id]");
    if (!act || !holder) return;
    const id = holder.dataset.id, cur = state.tracking[id] || {};
    if (act === "status") {
      el.dataset.s = el.value;
      saveTracking(id, statusPatch(cur, el.value)).then((ok) => { if (ok) toast(`Status set to ${el.value}`); });
    } else if (["applied_on", "contact", "notes", "next_action", "follow_up_on"].includes(act)) {
      const notApplied = TO_APPLY.has(cur.status || "New");
      if ((act === "applied_on" && notApplied ? "" : cur[act] || "") === el.value) return;
      const patch = { [act]: el.value };
      if (act === "applied_on" && el.value && notApplied) patch.status = "Applied";   // an applied date means she applied
      saveTracking(id, patch).then((ok) => { if (ok) toast(patch.status ? "Saved. Status set to Applied." : "Saved"); });
    }
  });
  document.addEventListener("focusout", () => { setTimeout(() => { if (state.pendingRender) render(); }, 0); });
  $("#q").addEventListener("input", (e) => { state.f.q = e.target.value; render(); });
  $("#showClosed").addEventListener("change", (e) => { state.f.showClosed = e.target.checked; render(); });
  $("#sort").addEventListener("change", (e) => { state.f.sort = e.target.value; render(); });
  $("#srcFilter").addEventListener("change", (e) => { state.f.src = e.target.value; render(); });
  window.addEventListener("hashchange", () => { const v = location.hash.slice(1); if (VIEWS[v] && v !== state.view) setView(v); });

  // drag and drop on the pipeline board (desktop); the status select covers touch devices
  let dragId = null;
  $("#board").addEventListener("dragstart", (ev) => { const card = ev.target.closest(".kcard"); if (!card || !state.writable) return; dragId = card.dataset.id; ev.dataTransfer.effectAllowed = "move"; ev.dataTransfer.setData("text/plain", dragId); card.classList.add("dragging"); });
  $("#board").addEventListener("dragend", (ev) => { const card = ev.target.closest(".kcard"); if (card) card.classList.remove("dragging"); $$(".col.drop").forEach((c) => c.classList.remove("drop")); });
  $("#board").addEventListener("dragover", (ev) => { const col = ev.target.closest(".col"); if (!col || !dragId) return; ev.preventDefault(); $$(".col.drop").forEach((c) => c !== col && c.classList.remove("drop")); col.classList.add("drop"); });
  $("#board").addEventListener("drop", (ev) => {
    const col = ev.target.closest(".col");
    $$(".col.drop").forEach((c) => c.classList.remove("drop"));
    if (!col || !dragId) return;
    ev.preventDefault();
    const id = dragId; dragId = null;
    const status = col.dataset.status, cur = state.tracking[id] || {};
    if ((cur.status || "New") === status) return;
    saveTracking(id, statusPatch(cur, status)).then((ok) => ok && toast(`Moved to ${status}`));
  });

  // ---------------------------------------------------------------- settings panel
  function fillSettings() {
    $("#s-owner").value = cfg.owner; $("#s-repo").value = cfg.repo; $("#s-branch").value = cfg.branch;
    $("#s-token").value = cfg.token ? "••••••••" : "";
    $("#connLine").innerHTML = cfg.owner && cfg.repo ? `Data: <b>${esc(cfg.owner)}/${esc(cfg.repo)}</b> on <b>${esc(cfg.branch)}</b> · ${cfg.token ? "token saved on this device" : "no token (read-only)"}` : "Not connected yet.";
  }
  function openSettings() { fillSettings(); $("#settingsPanel").hidden = false; $("#s-owner").focus(); }
  $("#settingsBtn").addEventListener("click", () => { const p = $("#settingsPanel"); if (p.hidden) openSettings(); else p.hidden = true; });
  $("#noticeBtn").addEventListener("click", openSettings);
  $("#settingsBtn2").addEventListener("click", () => { const p = $("#settingsPanel"); if (p.hidden) openSettings(); else p.hidden = true; });
  $("#excelBtn2").addEventListener("click", () => $("#excelBtn").click());
  $("#settingsCancel").addEventListener("click", () => { $("#settingsPanel").hidden = true; $("#settingsMsg").textContent = ""; });
  $("#forgetBtn").addEventListener("click", () => { cfg = Object.assign({}, cfg, { token: "" }); storeSettings(cfg); fillSettings(); state.writable = false; $("#settingsMsg").textContent = "Token removed from this device."; refresh(); });
  $("#settingsForm").addEventListener("submit", (ev) => {
    ev.preventDefault();
    const t = $("#s-token").value.trim();
    const next = { owner: $("#s-owner").value.trim(), repo: $("#s-repo").value.trim(), branch: $("#s-branch").value.trim() || "main", token: t === "••••••••" ? cfg.token : t };
    if (!next.owner || !next.repo) { $("#settingsMsg").textContent = "Add the GitHub owner and the repository name."; return; }
    if (next.token && !/^(github_pat_|ghp_)[A-Za-z0-9_]+$/.test(next.token)) { $("#settingsMsg").textContent = "That doesn't look like a GitHub token (it starts with github_pat_ or ghp_)."; return; }
    cfg = next;
    $("#settingsMsg").textContent = storeSettings(cfg) ? "Saved." : "This browser blocks saving settings (private mode?). They'll last until you close the tab.";
    state.loaded = false; state.loadError = null; state.commits = null;
    fillSettings();
    refresh().then(() => { if (state.loaded) { $("#settingsPanel").hidden = true; toast(state.writable ? "Connected. Changes save to GitHub." : "Connected (read-only)."); } });
  });

  // ---------------------------------------------------------------- add a role
  const KEEP = new Set(["jid", "jl", "gh_jid", "id", "jobid", "job_id"]);
  function canonical(u) {
    const url = new URL(u.trim());
    const q = new URLSearchParams();
    for (const [k, v] of url.searchParams) if (KEEP.has(k.toLowerCase())) q.append(k, v);
    const path = url.pathname.replace(/\/+$/, "") || "/";
    const qs = q.toString();
    return `${url.protocol.replace(":", "").toLowerCase()}://${url.host.toLowerCase().replace(/^www\./, "")}${path}${qs ? "?" + qs : ""}`;
  }
  async function sha12(text) {
    const buf = await crypto.subtle.digest("SHA-1", new TextEncoder().encode(text));
    return Array.from(new Uint8Array(buf)).map((b) => b.toString(16).padStart(2, "0")).join("").slice(0, 12);
  }
  $("#addBtn").addEventListener("click", () => { const p = $("#addPanel"); p.hidden = !p.hidden; if (!p.hidden) $("#f-title").focus(); });
  $("#addCancel").addEventListener("click", () => { $("#addPanel").hidden = true; $("#addMsg").textContent = ""; });
  $("#addForm").addEventListener("submit", async (ev) => {
    ev.preventDefault();
    const msg = $("#addMsg");
    const title = $("#f-title").value.trim(), company = $("#f-company").value.trim(), url = $("#f-url").value.trim();
    if (!title || !company) { msg.textContent = "Add at least a title and a company."; return; }
    if (url && !/^https?:\/\/\S+\.\S+/i.test(url)) { msg.textContent = "The link should start with https://"; return; }
    if (!state.writable) { msg.textContent = "Add a GitHub token in Settings first."; return; }
    msg.textContent = "Saving…";
    try {
      const canon = url ? canonical(url) : "";
      const id = await sha12(url ? canon : `lead:${company.toLowerCase()}|${title.toLowerCase()}`);
      const today = localISO();
      const fit = $("#f-fit").value ? Number($("#f-fit").value) : null;
      const doc = {
        id, title, company, url, canonical_url: canon,
        location: $("#f-location").value.trim() || "Not given", mode: $("#f-mode").value, type: $("#f-type").value,
        pay: $("#f-pay").value.trim() || "Not listed", posted: "Not shown", found_on: today, found_run: "", source: SRC_LABEL[$("#f-source").value] || "Added by you",
        source_type: $("#f-source").value, req_id: $("#f-req").value.trim(), fit, tags: [], why: "", watch: "", auth: "Not stated",
        posting_status: "Open", last_checked: url ? today : "", added_by: "manual", rank: 5000,
      };
      const { skipped } = await commitChange("jobs", () => ({ schema_version: 1, jobs: [] }), (d) => {
        d.jobs = Array.isArray(d.jobs) ? d.jobs : [];
        if (d.jobs.some((j) => j.id === id)) return false;
        d.jobs.push(doc); d.updated_at = new Date().toISOString();
      }, `Tracker: add ${title.slice(0, 60)} (${company.slice(0, 40)})`);
      if (skipped) { msg.textContent = "That role is already in the tracker."; return; }
      state.jobs.push(doc);
      const nextAction = $("#f-next").value.trim();
      if (nextAction) await saveTracking(id, { next_action: nextAction, status: "Shortlisted" }, "next action");
      $("#addForm").reset(); msg.textContent = ""; $("#addPanel").hidden = true; render(); toast("Role added");
    } catch (e) {
      msg.textContent = e && (e.status === 401 || e.status === 403) ? "GitHub refused the save. Check the token in Settings." : "That role didn't save. Try again.";
    }
  });

  // ---------------------------------------------------------------- contact form
  function openContactForm(c) {
    state.editingContact = c ? c.id : null;
    $("#contactTitle").textContent = c ? `Edit ${c.name}` : "Add contact";
    $("#c-name").value = c ? c.name || "" : ""; $("#c-company").value = c ? c.company || "" : "";
    $("#c-rel").value = c ? c.relationship || "Recruiter" : "Recruiter"; $("#c-role").value = c ? c.related_role || "" : "";
    $("#c-status").value = c ? c.status || "Active" : "Active"; $("#c-reach").value = c ? c.reach || "" : "";
    $("#c-last").value = c ? c.last_contact || "" : ""; $("#c-next").value = c ? c.next_follow_up || "" : "";
    $("#c-notes").value = c ? c.notes || "" : ""; $("#contactMsg").textContent = "";
    $("#contactPanel").hidden = false; $("#c-name").focus();
  }
  $("#contactAddBtn").addEventListener("click", () => { if (!state.writable) { toast("Add a GitHub token in Settings to save contacts."); return; } openContactForm(null); });
  $("#contactCancel").addEventListener("click", () => { $("#contactPanel").hidden = true; state.editingContact = null; });
  $("#contactForm").addEventListener("submit", (ev) => {
    ev.preventDefault();
    const name = $("#c-name").value.trim();
    if (!name) { $("#contactMsg").textContent = "Add a name."; return; }
    const existing = state.contacts.find((c) => c.id === state.editingContact);
    const rnd = Array.from(crypto.getRandomValues(new Uint8Array(5))).map((b) => b.toString(16).padStart(2, "0")).join("");
    const contact = Object.assign({}, existing || { id: `c-${rnd}`, created_on: localISO() }, {
      name, company: $("#c-company").value.trim(), relationship: $("#c-rel").value, related_role: $("#c-role").value.trim(),
      status: $("#c-status").value, reach: $("#c-reach").value.trim(), last_contact: $("#c-last").value, next_follow_up: $("#c-next").value, notes: $("#c-notes").value.trim(),
    });
    $("#contactPanel").hidden = true; state.editingContact = null;
    saveContact(contact, existing ? "edited" : "added").then((ok) => ok && toast(existing ? "Contact saved" : "Contact added"));
  });

  // ---------------------------------------------------------------- re-check postings (GitHub Actions workflow in the data repo)
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  function recheckCounts() {
    const today = localISO(), all = merged();
    return { stale: all.filter((j) => isStale(j, today)).length, open: all.filter((j) => !isPassed(j) && /^https?:/i.test(j.url || "")).length };
  }
  function renderRecheck() {
    const c = recheckCounts();
    $("#rcStale").textContent = `Re-check stale (${c.stale})`;
    $("#rcAll").textContent = `Re-check all open (${c.open})`;
    $("#rcStale").disabled = state.recheckBusy || !c.stale;
    $("#rcAll").disabled = state.recheckBusy || !c.open;
    $("#recheckBtn").textContent = state.recheckBusy ? "Re-checking…" : "Re-check postings";
    const r = state.reverify;
    $("#rcLast").innerHTML = r && r.last_run_at ? `Last re-check: <b>${esc(r.last_run_label || fmtCT(r.last_run_at))}</b>${r.last_counts ? ` · ${r.last_counts.checked} checked, ${r.last_counts.open} open, ${r.last_counts.closed} closed, ${r.last_counts.unknown} to check yourself` : ""}` : "No re-checks yet. The twice-daily search also re-checks the 8 oldest roles every morning.";
  }
  const setRc = (msg) => { $("#rcStatus").textContent = msg; };
  async function workflowRuns(perPage) {
    const r = await fetch(`${repoUrl()}/actions/workflows/${WORKFLOW}/runs?per_page=${perPage}`, { headers: headers(), cache: "no-store" });
    if (!r.ok) throw httpError(r.status);
    return (await r.json()).workflow_runs || [];
  }
  async function startRecheck(scope, ids) {
    if (!state.writable) { toast("Add a GitHub token in Settings first."); return; }
    if (state.recheckBusy) { toast("A re-check is already running."); return; }
    state.recheckBusy = true; $("#recheckPanel").hidden = false; renderRecheck();
    setRc(ids && ids.length ? "Starting a check of that posting…" : "Starting…");
    try {
      const before = ((await workflowRuns(1))[0] || {}).id || 0;
      const d = await fetch(`${repoUrl()}/actions/workflows/${WORKFLOW}/dispatches`, {
        method: "POST", headers: Object.assign(headers(), { "Content-Type": "application/json" }),
        body: JSON.stringify({ ref: cfg.branch, inputs: { scope, ids: (ids || []).join(",") } }),
      });
      if (d.status !== 204) throw httpError(d.status);
      const t0 = Date.now();
      let run = null;
      while (Date.now() - t0 < 6 * 60000) {
        await sleep(4000);
        run = (await workflowRuns(5)).find((x) => x.id > before) || null;
        const secs = Math.round((Date.now() - t0) / 1000);
        if (run && run.status === "completed") break;
        setRc(run ? `Checking postings… ${secs}s` : `Waiting for GitHub to start the check… ${secs}s`);
      }
      if (!run || run.status !== "completed") { setRc("Still running on GitHub. Results will appear here when it finishes; reload in a minute."); return; }
      if (run.conclusion !== "success") { setRc("The check didn't finish. Open the Actions tab of the data repository to see why."); return; }
      await refresh();
      const c = state.reverify && state.reverify.last_counts;
      const msg = c ? `Re-checked ${c.checked}: ${c.open} still open, ${c.closed} closed, ${c.unknown} to check yourself.` : "Re-check finished.";
      setRc(msg); toast(msg);
    } catch (e) {
      if (e.status === 403) setRc("Your token can't start re-checks yet. On GitHub, edit the token (Settings → Developer settings → Fine-grained tokens) and set Actions to Read and write for the data repository.");
      else if (e.status === 404) setRc("The re-check workflow isn't in the data repository yet, or the token can't see it.");
      else if (e.status === 422) setRc("GitHub rejected the request. Check the branch in Settings.");
      else setRc("Couldn't start the check. Check your connection and try again.");
    } finally {
      state.recheckBusy = false; renderRecheck();
    }
  }
  $("#recheckBtn").addEventListener("click", () => { const p = $("#recheckPanel"); p.hidden = !p.hidden; renderRecheck(); });
  $("#rcClose").addEventListener("click", () => { $("#recheckPanel").hidden = true; });
  $("#rcStale").addEventListener("click", () => startRecheck("stale", []));
  $("#rcAll").addEventListener("click", () => startRecheck("all", []));

  // ---------------------------------------------------------------- Excel
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
      const wb = buildWorkbook(ExcelJS, sortRows(merged(), "match"), today, state.profile, state.contacts);
      const buf = await wb.xlsx.writeBuffer();
      const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob); a.download = `Sumana_Job_Tracker_${today}.xlsx`;
      document.body.appendChild(a); a.click();
      setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
      toast("Excel downloaded");
    } catch (_) { toast("The Excel file couldn't be built. Check your connection and try again."); }
    finally { btn.disabled = false; btn.textContent = "Download Excel"; }
  });

  // ---------------------------------------------------------------- boot
  render();
  refresh();
  setInterval(() => { $("#viewSub").textContent = `${longToday()} · ${nextRunText()}`; }, 60000);
  setInterval(() => { if (document.visibilityState === "visible" && pendingWrites === 0) refresh(); }, 5 * 60000);
  document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible" && pendingWrites === 0) refresh(); });
})();

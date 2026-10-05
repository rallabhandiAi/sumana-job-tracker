# Sumana's Job Tracker

A small static website for Sumana's SoC design-verification job search. It lists verified openings with a match score, pay, work mode and a link to each posting, tracks every application through a status pipeline (New → Shortlisted → Applied → Recruiter screen → Interviewing → Offer), keeps notes and recruiter contacts, flags follow-ups 7 days after applying, and exports everything to Excel.

It runs on GitHub Pages.

> **This deployment:** the site lives in this public repo, and the data lives in the private repo `rallabhandiAi/sumana-job-tracker-data` (set in `assets/config.js`). Each device needs the token from step 4 once.
 The data is plain JSON in a GitHub repository, and status changes are saved as small commits through the GitHub API.

## What's in the repo

```
index.html            the app
assets/app.js         interface + GitHub data layer
assets/export.js      Excel export (ExcelJS)
assets/styles.css     styles (light and dark)
assets/config.js      optional defaults: owner, data repo, branch (no secrets)
assets/favicon.svg
data/jobs.json        roles (the twice-daily search adds to this file)
data/tracking.json    statuses, applied dates, contacts, notes (the site writes this file)
data/meta.json        search log (Activity page)
data/contacts.json    recruiters and referrals (Contacts page)
data/reverify.json    latest Re-check postings results
data/profile.json     search profile used for the Excel export (keep it in the private data repo)
resume/               resume files (Word, PDF) for the Resume panel (private data repo only)
.nojekyll             serve files as-is (optional)
```

## Setup (about 5 minutes)

### 1. Decide where the data lives

| Setup | Repos | Who can see statuses and notes |
|---|---|---|
| **Simple** | one public repo `sumana-job-tracker` with everything | anyone who finds the repo or the site |
| **Private data** (recommended) | public `sumana-job-tracker` (site only) + private `sumana-job-tracker-data` (only the `data/` folder) | only people with your token |

GitHub Pages on a free account needs the site repo to be public, and a Pages site is public even when its repo is private. That's why private data goes in a separate private repo that the site reads through the API.

For **private data**: put the `data/` folder in the private repo, delete `data/` from the site repo, and set `owner` and `dataRepo` in `assets/config.js`:

```js
window.TRACKER_CONFIG = { owner: "your-username", dataRepo: "sumana-job-tracker-data", branch: "main" };
```

### 2. Upload the files

On github.com: **New repository** → name it → create it without a README → **uploading an existing file** → drag in the contents of this folder → **Commit changes**.

Or with git:

```bash
git init && git add . && git commit -m "Job tracker"
git branch -M main
git remote add origin https://github.com/<your-username>/sumana-job-tracker.git
git push -u origin main
```

### 3. Turn on GitHub Pages

Repo → **Settings** → **Pages** → Build and deployment → Source: **Deploy from a branch** → Branch: `main`, folder `/ (root)` → **Save**. After a minute or two the site is live at `https://<your-username>.github.io/sumana-job-tracker/`.

### 4. Create a token for saving changes

GitHub → **Settings** → **Developer settings** → **Personal access tokens** → **Fine-grained tokens** → **Generate new token**:

- Repository access: **Only select repositories** → the repo that holds `data/`
- Repository permissions → **Contents: Read and write** and **Actions: Read and write** (Actions lets the Re-check postings button start its check)
- Expiration: your choice (for example 90 days)

Copy the token. It starts with `github_pat_`.

### 5. Connect each device

Open the site → **Settings** → check the owner, data repository and branch → paste the token → **Save and reload**. Do this once per device (laptop, phone). The token stays in that browser only. Without a token, a public repo opens read-only.

## Daily use

| Page | What it's for |
|---|---|
| **Today** | Counts at the top, then three lists: Apply next, Follow up (roles and contacts) and Re-check before applying, each with one-tap buttons |
| **Opportunities** | Every role with match score, source (official / job board / vendor / recruiter), next action, follow-up date and notes |
| **Pipeline** | Ready to apply → Applied → Recruiter screen → Interviewing → Offer. Drag cards on a laptop, or use the status dropdown |
| **Contacts** | Recruiters, referrals and hiring managers with follow-up dates (these show up on Today) |
| **Activity** | Every saved change and search run, from the data repo's commit history |

- Choosing **Applied** fills in today's date; the role shows **Follow up** 7 days later.
- **Re-check postings** opens each stale posting on GitHub's servers and marks gone or expired ones as Closed. Sites that block automatic checks are listed for you to open yourself.
- **Download Excel** builds a workbook (Dashboard, Jobs, Contacts, Search Profile) from the current data.
- **Resume** lists the files in the data repo's `resume` folder and downloads them with your token. Keep resumes in the private data repo only, never in this public site repo.
- **Add a role** saves one found elsewhere (ChatGPT, LinkedIn, a recruiter call). The link is optional for recruiter leads.
- Every change is a commit in the data repo, so you get a full history for free.

## How new roles arrive

A Claude scheduled task runs at 10:30 AM and 8:30 PM Central. It searches job boards, opens each posting to confirm it's still live, scores the match, appends new roles to `data/jobs.json`, and logs the run in `data/meta.json`. It never edits `data/tracking.json`. To let it write here, connect GitHub to Claude (claude.ai → Settings → Connectors → GitHub) with access to the repo that holds `data/`.

## Data format

`data/jobs.json`

```json
{ "schema_version": 1, "updated_at": "2026-10-05T16:56:01Z", "jobs": [ { "id": "13cd0a2348aa", "title": "…" } ] }
```

Each role: `id`, `title`, `company`, `location`, `mode` (Remote | Hybrid | On-site | Confirm), `type`, `pay`, `posted`, `source`, `url`, `canonical_url`, `fit` (3–5), `tags`, `why`, `watch`, `auth`, `found_on`, `found_run`, `posting_status` (Open | Closed), `last_checked`, `added_by`, `rank`.
`id` is the first 12 hex characters of the SHA-1 of the canonical posting URL.

`data/tracking.json`

```json
{ "schema_version": 1, "tracking": { "13cd0a2348aa": { "status": "Applied", "applied_on": "2026-10-06", "contact": "", "notes": "", "updated_at": "…" } } }
```

Statuses: New, Shortlisted, Applied, Recruiter screen, Interviewing, Offer, Rejected, Not a fit, Closed.

## Preview locally

```bash
python3 -m http.server 8000
```

Then open http://localhost:8000. Opening `index.html` straight from disk won't load the data files.

## Privacy and security

- The token lives in the browser's local storage for this site only. Use a fine-grained token limited to the one data repo with Contents access only, and revoke it on GitHub anytime.
- The page sets a strict Content-Security-Policy and loads ExcelJS from jsDelivr pinned to an exact version with an integrity hash.
- `noindex` keeps search engines from listing the site, but a public repo is still visible on GitHub.

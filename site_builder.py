"""
Publishes the tender finder results as a static web page (for GitHub Pages),
tracks which tenders are new since the last run, and emails new matches.

Email settings come from environment variables (GitHub repository secrets):
  SMTP_HOST, SMTP_PORT (587 or 465), SMTP_USER, SMTP_PASS,
  ALERT_TO (comma-separated), ALERT_FROM (optional), SITE_URL (optional)
"""

import html
import json
import os
import smtplib
import ssl
from datetime import date, datetime, timedelta
from email.message import EmailMessage
from pathlib import Path

PBKDF2_ITERATIONS = 250_000


def _password():
    return os.environ.get("SITE_PASSWORD") or None


def _encrypt(data: bytes, password: str) -> bytes:
    """salt(16) | iv(12) | AES-256-GCM ciphertext+tag. Matches the WebCrypto code in the page."""
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    salt, iv = os.urandom(16), os.urandom(12)
    key = PBKDF2HMAC(hashes.SHA256(), 32, salt, PBKDF2_ITERATIONS).derive(password.encode())
    return salt + iv + AESGCM(key).encrypt(iv, data, None)


def _decrypt(blob: bytes, password: str) -> bytes:
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    salt, iv, ct = blob[:16], blob[16:28], blob[28:]
    key = PBKDF2HMAC(hashes.SHA256(), 32, salt, PBKDF2_ITERATIONS).derive(password.encode())
    return AESGCM(key).decrypt(iv, ct, None)


def _load_payload(site_dir):
    """Previous run's data: encrypted file if a password is set, else plain JSON."""
    d = Path(site_dir)
    pw = _password()
    if pw and (d / "data.enc").exists():
        try:
            return json.loads(_decrypt((d / "data.enc").read_bytes(), pw))
        except Exception as e:
            print(f"Could not decrypt previous data ({e.__class__.__name__}); starting fresh")
            return {}
    if (d / "data.json").exists():
        try:
            return json.loads((d / "data.json").read_text())
        except ValueError:
            return {}
    return {}


def _key(r):
    return r.get("ocid") or r.get("url") or (r.get("buyer", "") + r.get("title", ""))


def _jsonable(r):
    return {k: (v.isoformat() if isinstance(v, datetime) else v) for k, v in r.items()}


def _parse(v):
    try:
        return datetime.fromisoformat(v) if isinstance(v, str) and len(v) >= 10 else v
    except ValueError:
        return v


# ---------------------------------------------------------------- state

def mark_new(site_dir, opps):
    """Tag each opportunity with first_seen / is_new. Returns the list of new ones."""
    path = Path(site_dir) / "state.json"
    first_run = not path.exists()
    state = json.loads(path.read_text()) if path.exists() else {}
    today = date.today().isoformat()
    new = []
    for r in opps:
        k = _key(r)
        added = k not in state
        if added:
            state[k] = today
            if not first_run:
                new.append(r)
        r["first_seen"] = state[k]
        r["is_new"] = added and not first_run
    cutoff = (date.today() - timedelta(days=365)).isoformat()
    state = {k: v for k, v in state.items() if v >= cutoff}
    Path(site_dir).mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(state, indent=0))
    if first_run:  # first ever run: alert with the strongest matches instead of everything
        open_ = [r for r in opps if not r.get("deadline") or r["deadline"] >= datetime.now()]
        new = sorted(open_, key=lambda r: -r.get("score", 0))[:15]
        for r in new:
            r["_first_run"] = True
    return new


def load_previous_opps(site_dir):
    data = _load_payload(site_dir)
    return [{k: _parse(v) for k, v in o.items()} for o in data.get("opps", [])]


def merge_opps(previous, fresh, keep_days=45):
    """Fresh notices replace older copies; earlier finds stay while still relevant."""
    now = datetime.now()
    merged = {}
    for r in previous:
        dl = r.get("deadline")
        seen = r.get("first_seen")
        try:
            seen_dt = datetime.fromisoformat(seen) if isinstance(seen, str) else None
        except ValueError:
            seen_dt = None
        if isinstance(dl, datetime):
            if dl < now - timedelta(days=1):
                continue          # closed: drop it
        elif seen_dt and seen_dt < now - timedelta(days=max(keep_days, 90)):
            continue              # no deadline and old: drop it
        r.pop("is_new", None)
        merged[_key(r)] = r
    for r in fresh:
        merged[_key(r)] = r
    return list(merged.values())


def load_previous_awards(site_dir):
    data = _load_payload(site_dir)
    return [{k: _parse(v) for k, v in a.items()} for a in data.get("awards", [])]


# ---------------------------------------------------------------- site

def write_site(site_dir, opps, awards, new_items, xlsx_path=None):
    d = Path(site_dir)
    d.mkdir(parents=True, exist_ok=True)
    seen, uniq = set(), []
    for a in awards:  # the same award can appear in two scan windows
        k = _key(a)
        if k not in seen:
            seen.add(k)
            uniq.append(a)
    awards = uniq
    payload = {
        "generated": datetime.now().isoformat(timespec="minutes"),
        "opps": [_jsonable(r) for r in opps],
        "awards": [_jsonable(r) for r in awards],
        "new_count": len([r for r in opps if r.get("is_new")]),
    }
    raw = json.dumps(payload, default=str)
    pw = _password()
    if pw:
        # Everything published is encrypted; remove any plain copies from earlier runs.
        (d / "data.enc").write_bytes(_encrypt(raw.encode(), pw))
        xlsx = Path(xlsx_path) if xlsx_path else d / "EPL_Tenders_latest.xlsx"
        if xlsx.exists():
            (d / "tenders.xlsx.enc").write_bytes(_encrypt(xlsx.read_bytes(), pw))
            xlsx.unlink()
        for old in ("data.json", "EPL_Tenders_latest.xlsx"):
            if (d / old).exists():
                (d / old).unlink()
        html_page = PAGE.replace("__DATA__", "null").replace("__ITER__", str(PBKDF2_ITERATIONS))
    else:
        (d / "data.json").write_text(raw)
        html_page = PAGE.replace("__DATA__", raw.replace("</", "<\\/")).replace("__ITER__", "0")
    (d / "index.html").write_text(html_page, encoding="utf-8")
    (d / ".nojekyll").write_text("")
    print(f"Site written to {d}/index.html")


# ---------------------------------------------------------------- email

def send_alert(new_items):
    host, to = os.environ.get("SMTP_HOST"), os.environ.get("ALERT_TO")
    if not host or not to:
        print("Email skipped: SMTP_HOST / ALERT_TO not set")
        return
    if not new_items:
        print("Email skipped: no new matches")
        return
    port = int(os.environ.get("SMTP_PORT", "587"))
    user, pw = os.environ.get("SMTP_USER"), os.environ.get("SMTP_PASS")
    site = os.environ.get("SITE_URL", "")
    first = any(r.get("_first_run") for r in new_items)
    items = sorted(new_items, key=lambda r: (r.get("deadline") or datetime.max))

    subject = (f"EPL tender finder is live: top {len(items)} matches" if first
               else f"EPL tenders: {len(items)} new match{'es' if len(items) != 1 else ''}")

    def line(r):
        dl = r["deadline"].strftime("%d %b %Y %H:%M") if r.get("deadline") else "no deadline given"
        val = f"£{r['value']:,.0f}" if isinstance(r.get("value"), (int, float)) else "value not stated"
        return r, dl, val

    text = [subject, ""]
    rows = []
    for r, dl, val in map(line, items):
        stage = " (pipeline notice)" if r.get("stage") == "planning" else ""
        text += [f"- {r['title']}{stage}", f"  {r['buyer']} | {val} | closes {dl}",
                 f"  {r.get('email') or ''} {r.get('phone') or ''}".rstrip(), f"  {r['url']}", ""]
        rows.append(
            f"<tr><td style='padding:8px 10px;border-bottom:1px solid #ddd'>"
            f"<a href='{html.escape(r['url'])}' style='color:#2a6a6e;font-weight:600'>{html.escape(r['title'])}</a>"
            f"{' <em>(pipeline notice)</em>' if stage else ''}<br>"
            f"<span style='color:#555'>{html.escape(r['buyer'])}</span><br>"
            f"<span style='color:#555'>{html.escape(r.get('category') or '')}</span></td>"
            f"<td style='padding:8px 10px;border-bottom:1px solid #ddd;white-space:nowrap'>{html.escape(val)}<br>closes {html.escape(dl)}</td>"
            f"<td style='padding:8px 10px;border-bottom:1px solid #ddd'>{html.escape(r.get('contact_name') or '')}<br>"
            f"{html.escape(r.get('email') or '')}<br>{html.escape(r.get('phone') or '')}</td></tr>")
    if site:
        text.append(f"All tenders: {site}")
    body_html = (f"<div style='font-family:Arial,sans-serif;font-size:14px;color:#16252d'>"
                 f"<h2 style='margin:0 0 12px'>{html.escape(subject)}</h2>"
                 f"<table style='border-collapse:collapse;width:100%'>{''.join(rows)}</table>"
                 + (f"<p><a href='{html.escape(site)}'>Open the tender finder</a></p>" if site else "")
                 + "<p style='color:#777;font-size:12px'>While a tender is live, contact the buyer only "
                   "through the portal's clarification route.</p></div>")

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = os.environ.get("ALERT_FROM") or user
    msg["To"] = to
    msg.set_content("\n".join(text))
    msg.add_alternative(body_html, subtype="html")

    ctx = ssl.create_default_context()
    if port == 465:
        with smtplib.SMTP_SSL(host, port, context=ctx, timeout=60) as s:
            if user:
                s.login(user, pw)
            s.send_message(msg)
    else:
        with smtplib.SMTP(host, port, timeout=60) as s:
            s.starttls(context=ctx)
            if user:
                s.login(user, pw)
            s.send_message(msg)
    print(f"Emailed {len(items)} matches to {to}")


# ---------------------------------------------------------------- page

PAGE = r"""<!DOCTYPE html>
<html lang="en-GB">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<meta name="robots" content="noindex, nofollow">
<title>EPL Tender Finder</title>
<link rel="icon" href="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><text y='.9em' font-size='90'>⚡</text></svg>">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Archivo:wdth,wght@75..125,400..800&display=swap" rel="stylesheet">
<style>
:root{--ink:#16252d;--paper:#eef1ec;--panel:#fff;--line:#cfd6d0;--muted:#5f6f75;--grid:#2a6a6e;--signal:#e9b623;--signal-ink:#3b2c00;--hot:#b4473a;--warn:#9a6f00;--soft:#e3e8e3;
  box-sizing:border-box;padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){--ink:#e6ece8;--paper:#101a1f;--panel:#18242a;--line:#2c3b42;--muted:#93a3a8;--grid:#5fb3b6;--hot:#e07a6c;--warn:#e9b623;--soft:#1f2e35}}
:root[data-theme="dark"]{--ink:#e6ece8;--paper:#101a1f;--panel:#18242a;--line:#2c3b42;--muted:#93a3a8;--grid:#5fb3b6;--hot:#e07a6c;--warn:#e9b623;--soft:#1f2e35}
*,*::before,*::after{box-sizing:inherit}
[hidden]{display:none!important}
html{scroll-padding-top:env(safe-area-inset-top,0px)}
body{margin:0;background:var(--paper);color:var(--ink);font-family:"Archivo",system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;font-size:15px;line-height:1.45;font-variant-numeric:tabular-nums}
a{color:var(--grid)}
:focus-visible{outline:2px solid var(--grid);outline-offset:2px}
.wrap{max-width:1180px;margin:0 auto;padding:20px 20px 60px}
header{display:flex;flex-wrap:wrap;justify-content:space-between;align-items:flex-end;gap:14px;padding-bottom:16px;border-bottom:3px solid var(--ink)}
h1{margin:0;font-weight:800;font-stretch:78%;font-size:clamp(30px,4.6vw,48px);line-height:.95}
.updated{margin:8px 0 0;color:var(--muted)}
.btn{display:inline-block;border:1.5px solid var(--ink);background:var(--ink);color:var(--paper);padding:8px 14px;border-radius:3px;font-weight:600;text-decoration:none}
.tabs{display:flex;flex-wrap:wrap;margin:22px 0 0}
.tabs button{font:inherit;border:0;background:none;padding:8px 2px;margin-right:20px;font-weight:700;font-size:17px;cursor:pointer;color:var(--muted);border-bottom:3px solid transparent}
.tabs button[aria-selected="true"]{color:var(--ink);border-bottom-color:var(--signal)}
.tabs .n{font-weight:500}
.filters{display:flex;flex-wrap:wrap;gap:10px;align-items:center;padding:12px 0;border-bottom:1px solid var(--line)}
.filters input[type=search]{flex:1 1 240px;min-width:0;font:inherit;padding:8px 10px;border:1.5px solid var(--line);border-radius:3px;background:var(--panel);color:inherit}
.filters select{font:inherit;padding:8px;border:1.5px solid var(--line);border-radius:3px;background:var(--panel);color:inherit}
.filters label{display:flex;gap:6px;align-items:center;font-size:14px}
details{border-bottom:1px solid var(--line)}
summary{display:grid;grid-template-columns:72px minmax(0,1fr) auto;gap:14px;align-items:center;padding:12px 6px;cursor:pointer;list-style:none}
summary::-webkit-details-marker{display:none}
summary:hover{background:var(--soft)}
details[open] summary{background:var(--soft)}
.days{font-stretch:75%;font-weight:800;font-size:30px;line-height:.9;text-align:right}
.days small{display:block;font-size:11px;font-weight:600;font-stretch:100%;color:var(--muted);margin-top:3px}
.days.hot{color:var(--hot)}.days.warn{color:var(--warn)}.days.closed{color:var(--muted);font-size:15px}
.t{font-weight:650;line-height:1.25}
.meta{font-size:13px;color:var(--muted);margin-top:3px;display:flex;flex-wrap:wrap;gap:3px 12px}
.cat{color:var(--grid);font-weight:600}
.new{background:var(--signal);color:var(--signal-ink);font-size:12px;font-weight:700;padding:1px 7px;border-radius:3px}
.val{font-weight:700;white-space:nowrap;text-align:right}
.body{padding:4px 6px 18px 92px;display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:4px 28px;font-size:14px}
.body h4{margin:12px 0 4px;font-size:13px;color:var(--muted);font-weight:600}
.body p{margin:0;overflow-wrap:anywhere}
.body .wide{grid-column:1/-1}
.desc{white-space:pre-wrap}
.empty{padding:30px 6px;color:var(--muted)}
.tablewrap{overflow-x:auto}
table{border-collapse:collapse;width:100%;min-width:820px;font-size:14px}
th,td{text-align:left;padding:9px 8px;border-bottom:1px solid var(--line);vertical-align:top}
th{font-size:13px;color:var(--muted);font-weight:600;border-bottom:1.5px solid var(--ink)}
.note{font-size:13px;color:var(--muted)}
.gate{max-width:360px;margin:12vh auto 0;padding:28px 24px;background:var(--panel);border:1.5px solid var(--ink);border-radius:4px;display:flex;flex-direction:column;gap:10px}
.gate h1{font-size:36px}
.gate input[type=password]{font:inherit;padding:10px;border:1.5px solid var(--line);border-radius:3px;background:var(--paper);color:inherit}
.gate .btn{cursor:pointer;font:inherit;font-weight:600}
.remember{display:flex;gap:6px;align-items:center}
.err{color:var(--hot);margin:0;min-height:1.2em;font-size:14px}
.btn.ghost{background:transparent;color:var(--ink);cursor:pointer;font:inherit;font-weight:600}
@media (max-width:700px){summary{grid-template-columns:56px minmax(0,1fr)}.val{grid-column:2;text-align:left}.body{padding-left:6px}}
</style>
</head>
<body>
<div id="gate" hidden>
  <form id="gateForm" class="gate" autocomplete="on">
    <h1>Tender finder</h1>
    <p class="note">Ecologic Partners. Enter the team password to continue.</p>
    <label for="pw" class="note">Password</label>
    <input id="pw" type="password" autocomplete="current-password" required autofocus>
    <label class="note remember"><input type="checkbox" id="remember"> Stay signed in on this device</label>
    <button class="btn" type="submit" id="unlock">Unlock</button>
    <p class="err" id="gateErr" role="alert"></p>
  </form>
</div>
<div class="wrap" id="site" hidden>
<header>
  <div><h1>Tender finder</h1><p class="updated" id="updated"></p></div>
  <div style="display:flex;gap:8px;flex-wrap:wrap">
    <a class="btn" id="dl" href="EPL_Tenders_latest.xlsx" download>Download Excel</a>
    <button class="btn ghost" type="button" id="signout" hidden>Sign out</button>
  </div>
</header>
<div class="tabs" role="tablist">
  <button role="tab" data-v="live" aria-selected="true" type="button">Live tenders <span class="n" id="nLive"></span></button>
  <button role="tab" data-v="plan" aria-selected="false" type="button">Coming up <span class="n" id="nPlan"></span></button>
  <button role="tab" data-v="ren" aria-selected="false" type="button">Renewals <span class="n" id="nRen"></span></button>
</div>
<div class="filters">
  <input type="search" id="q" placeholder="Search title, buyer, region or keyword" aria-label="Search">
  <select id="cat" aria-label="Category"><option value="">All categories</option></select>
  <label id="newWrap"><input type="checkbox" id="newOnly"> New today only</label>
  <label id="renWrap" hidden>Ending within
    <select id="months"><option value="6">6 months</option><option value="12">12 months</option><option value="18" selected>18 months</option><option value="36">3 years</option><option value="9999">Any time</option></select></label>
</div>
<div id="out"></div>
<p class="note" style="margin-top:24px">Sources: Find a Tender, Contracts Finder, Public Contracts Scotland and Sell2Wales. While a tender is live, contact the buyer only through the portal's clarification route.</p>
</div>
<script>
const EMBEDDED = __DATA__;
const ITER = __ITER__;
const $ = id => document.getElementById(id);
let KEY = null;

const b64 = b => btoa(String.fromCharCode(...new Uint8Array(b)));
const unb64 = s => Uint8Array.from(atob(s), c => c.charCodeAt(0));
async function keyFor(pw, salt){
  const base = await crypto.subtle.importKey("raw", new TextEncoder().encode(pw), "PBKDF2", false, ["deriveBits"]);
  return crypto.subtle.deriveBits({name:"PBKDF2", hash:"SHA-256", salt, iterations: ITER}, base, 256);
}
async function decryptBlob(buf, rawKeyOrPw){
  const bytes = new Uint8Array(buf), salt = bytes.slice(0,16), iv = bytes.slice(16,28), ct = bytes.slice(28);
  const bits = typeof rawKeyOrPw === "string" ? await keyFor(rawKeyOrPw, salt) : null;
  const key = await crypto.subtle.importKey("raw", bits, "AES-GCM", false, ["decrypt"]);
  return { plain: await crypto.subtle.decrypt({name:"AES-GCM", iv}, key, ct), bits };
}
async function fetchBuf(name){
  const r = await fetch(name + "?v=" + Date.now(), {cache: "no-store"});
  if (!r.ok) throw new Error("missing " + name);
  return r.arrayBuffer();
}
const store = {
  get(){ try { return sessionStorage.getItem("epl_pw") || localStorage.getItem("epl_pw"); } catch { return null; } },
  set(pw, keep){ try { (keep ? localStorage : sessionStorage).setItem("epl_pw", pw); } catch {} },
  clear(){ try { sessionStorage.removeItem("epl_pw"); localStorage.removeItem("epl_pw"); } catch {} }
};
async function unlock(pw, keep){
  const { plain } = await decryptBlob(await fetchBuf("data.enc"), pw);
  KEY = pw;
  store.set(pw, keep);
  $("gate").hidden = true; $("site").hidden = false; $("signout").hidden = false;
  start(JSON.parse(new TextDecoder().decode(plain)));
}
function boot(){
  if (EMBEDDED) { $("site").hidden = false; start(EMBEDDED); return; }
  $("dl").removeAttribute("href");
  $("dl").addEventListener("click", async e => {
    e.preventDefault();
    try {
      const { plain } = await decryptBlob(await fetchBuf("tenders.xlsx.enc"), KEY);
      const url = URL.createObjectURL(new Blob([plain], {type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}));
      const a = document.createElement("a"); a.href = url; a.download = "EPL_Tenders_latest.xlsx"; a.click();
      setTimeout(() => URL.revokeObjectURL(url), 5000);
    } catch { alert("The Excel file couldn't be opened. Try again after the next scan."); }
  });
  $("signout").onclick = () => { store.clear(); location.reload(); };
  const saved = store.get();
  const showGate = msg => { $("gate").hidden = false; $("gateErr").textContent = msg || ""; $("pw").focus(); };
  $("gateForm").addEventListener("submit", async e => {
    e.preventDefault();
    $("unlock").disabled = true; $("gateErr").textContent = "";
    try { await unlock($("pw").value, $("remember").checked); }
    catch (err) { showGate(String(err.message).startsWith("missing") ? "The site is still being built. Try again in a few minutes." : "That password isn't right."); }
    finally { $("unlock").disabled = false; }
  });
  if (saved) unlock(saved, false).catch(() => { store.clear(); showGate(); });
  else showGate();
}

function start(DATA){
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const safe = u => /^https?:\/\//i.test(String(u||"").trim()) ? String(u).trim() : "";
const link = (u, label) => safe(u) ? `<a href="${esc(safe(u))}" target="_blank" rel="noopener">${esc(label || u)}</a>` : esc(u);
const money = v => typeof v === "number" ? "£" + Math.round(v).toLocaleString("en-GB") : "";
const dt = (v, time) => v ? new Date(v).toLocaleString("en-GB", time ? {day:"numeric",month:"short",year:"numeric",hour:"2-digit",minute:"2-digit"} : {day:"numeric",month:"short",year:"numeric"}) : "";
const daysLeft = v => v ? Math.floor((new Date(v) - Date.now()) / 86400000) : null;
let view = "live";

const opps = DATA.opps;
const live = opps.filter(r => r.stage !== "planning" && (!r.deadline || daysLeft(r.deadline) >= 0));
const plan = opps.filter(r => r.stage === "planning");
const ren = DATA.awards.map(a => ({...a, months: a.contract_end ? Math.round((new Date(a.contract_end) - Date.now()) / (86400000*30.4)) : null}));

$("updated").textContent = `Updated ${dt(DATA.generated, true)}. ${live.length} live tenders, ${DATA.new_count} new today.`;
$("nLive").textContent = live.length; $("nPlan").textContent = plan.length; $("nRen").textContent = ren.length;
[...new Set(opps.map(r => r.category).filter(Boolean))].sort().forEach(c => $("cat").insertAdjacentHTML("beforeend", `<option>${esc(c)}</option>`));

function match(r){
  const q = $("q").value.trim().toLowerCase(), c = $("cat").value;
  if (c && r.category !== c) return false;
  if (q && ![r.title, r.buyer, r.region, r.matched, r.category, r.suppliers].join(" ").toLowerCase().includes(q)) return false;
  return true;
}

function row(r){
  const d = daysLeft(r.deadline);
  const days = d === null ? `<div class="days closed">${r.stage === "planning" ? "Early" : "No date"}</div>`
    : d < 0 ? `<div class="days closed">Closed</div>`
    : `<div class="days ${d<=7?"hot":d<=14?"warn":""}">${d}<small>${d===1?"day":"days"} left</small></div>`;
  const docs = String(r.documents || "").split(/\s+/).filter(safe).slice(0, 8);
  const kv = (k, v) => v || v === 0 ? `<h4>${k}</h4><p>${v}</p>` : "";
  return `<details><summary>${days}
    <div><div class="t">${esc(r.title)}</div>
      <div class="meta">${r.is_new ? `<span class="new">New</span>` : ""}<span>${esc(r.buyer)}</span>${r.region ? `<span>${esc(r.region)}</span>` : ""}<span class="cat">${esc(r.category)}</span></div></div>
    <div class="val">${money(r.value)}</div></summary>
    <div class="body">
      <div>${kv("Deadline", esc(dt(r.deadline, true)))}${kv("Clarifications by", esc(dt(r.clarification_deadline, true)))}
        ${kv("Contract", [r.contract_start ? "from " + dt(r.contract_start) : "", r.duration_days ? Math.round(r.duration_days/30.4) + " months" : ""].filter(Boolean).join(", "))}
        ${kv("Procedure", esc(r.procedure))}${kv("Framework", esc(r.framework))}${kv("SME suitable", esc(r.sme))}</div>
      <div>${kv("Contact", [esc(r.contact_name), r.email ? `<a href="mailto:${esc(r.email)}">${esc(r.email)}</a>` : "", esc(r.phone)].filter(Boolean).join("<br>"))}
        ${kv("Address", esc(r.address))}${kv("Website", link(r.website))}${kv("Other contacts", esc(r.other_contacts))}</div>
      <div>${kv("Notice", link(r.url, "Open notice"))}${kv("Tender portal", link(r.portal))}
        ${docs.length ? kv("Documents", docs.map((u,i) => link(u, "Document " + (i+1))).join("<br>")) : ""}
        ${kv("Matched on", esc(r.matched))}${kv("First seen", esc(dt(r.first_seen)))}</div>
      ${r.award_criteria ? `<div class="wide">${kv("Award criteria", esc(r.award_criteria))}</div>` : ""}
      ${r.selection_criteria ? `<div class="wide">${kv("Selection criteria", esc(r.selection_criteria))}</div>` : ""}
      ${r.description ? `<div class="wide">${kv("Scope", `<span class="desc">${esc(r.description)}</span>`)}</div>` : ""}
    </div></details>`;
}

function render(){
  $("newWrap").hidden = view === "ren"; $("renWrap").hidden = view !== "ren";
  const out = $("out");
  if (view === "ren") {
    const m = +$("months").value;
    const rows = ren.filter(match).filter(r => r.months === null ? m === 9999 : r.months >= -1 && r.months <= m).sort((a,b) => (a.months ?? 1e9) - (b.months ?? 1e9));
    out.innerHTML = rows.length ? `<div class="tablewrap"><table><thead><tr><th>Ends</th><th>Contract</th><th>Buyer</th><th>Incumbent</th><th>Value</th><th>Contact</th></tr></thead><tbody>${rows.map(r => `<tr>
      <td><strong>${r.months === null ? "Unknown" : r.months <= 0 ? "Now" : r.months + " mo"}</strong><br><span class="note">${esc(dt(r.contract_end))}</span></td>
      <td>${link(r.url, r.title)}<br><span class="cat">${esc(r.category)}</span></td><td>${esc(r.buyer)}<br><span class="note">${esc(r.region)}</span></td>
      <td>${esc(r.suppliers)}</td><td>${money(r.award_value)}</td>
      <td>${esc(r.contact_name)}${r.email ? `<br><a href="mailto:${esc(r.email)}">${esc(r.email)}</a>` : ""}${r.phone ? "<br>" + esc(r.phone) : ""}</td></tr>`).join("")}</tbody></table></div>`
      : `<p class="empty">No contracts end in this window. Choose a longer range above. Award data refreshes every Monday.</p>`;
    return;
  }
  const set = (view === "live" ? live : plan).filter(match).filter(r => !$("newOnly").checked || r.is_new)
    .sort((a, b) => (a.deadline ? new Date(a.deadline) : Infinity) - (b.deadline ? new Date(b.deadline) : Infinity) || b.score - a.score);
  out.innerHTML = set.length ? set.map(row).join("") : `<p class="empty">Nothing matches. Clear the search or untick New today only.</p>`;
}
document.querySelectorAll(".tabs button").forEach(b => b.onclick = () => {
  view = b.dataset.v; document.querySelectorAll(".tabs button").forEach(x => x.setAttribute("aria-selected", x === b)); render(); });
["q","cat","newOnly","months"].forEach(id => $(id).addEventListener("input", render));
render();
}
boot();
</script>
</body>
</html>
"""

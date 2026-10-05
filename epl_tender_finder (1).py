#!/usr/bin/env python3
"""
EPL Tender Finder
-----------------
Pulls live UK public-sector tenders relevant to an energy consultancy / broker
from the two official government sources (no API key needed):

  * Find a Tender Service (FTS)  - above-threshold contracts
  * Contracts Finder (CF)        - below-threshold contracts (England)

It scores each notice for relevance (CPV codes + keywords), extracts buyer
contact details, deadlines, values, award/selection criteria, documents and
the e-tendering portal link, then writes a formatted Excel workbook:

  Opportunities   - live tenders and pipeline notices, ranked
  Buyer Contacts  - one row per buyer contact, de-duplicated
  Incumbents      - (with --awards) who won similar contracts and when they end
  About           - run settings and notes

Usage:
  pip install requests openpyxl
  python epl_tender_finder.py                       # last 30 days, live tenders
  python epl_tender_finder.py --days 60 --open-only
  python epl_tender_finder.py --region "North East" # or a NUTS code e.g. UKC
  python epl_tender_finder.py --awards --award-days 1095   # incumbent / renewal intel
"""

import argparse
import re
import sys
import time
from datetime import datetime, timedelta, timezone

import requests
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------
# CONFIG - tune these to EPL's service lines
# --------------------------------------------------------------------------

# Keyword groups -> category label. Matched as whole words, case-insensitive.
KEYWORD_GROUPS = {
    "Brokerage / Consultancy": [
        "energy broker", "energy brokerage", "utility broker", "utilities broker",
        "energy procurement", "utilities procurement", "utility procurement",
        "energy consultancy", "energy consultant", "utility consultancy",
        "utilities consultancy", "utilities management", "utility management",
        "energy management", "bill validation", "utility bill", "bureau service",
        "flexible procurement", "energy purchasing", "energy and utilities",
        "energy services", "meter data", "sub-metering", "sub metering",
    ],
    "Energy Supply": [
        "supply of electricity", "supply of gas", "electricity supply",
        "gas supply", "natural gas", "half hourly", "non half hourly",
        "water retail", "utilities supply",
    ],
    "Renewables / BESS / EV": [
        "solar pv", "photovoltaic", "solar panels", "battery storage", "bess",
        "battery energy storage", "ev charging", "electric vehicle charging",
        "chargepoint", "charge point", "charging infrastructure",
        "power purchase agreement", "ppa", "cppa", "private wire", "renewable energy",
    ],
    "Carbon / Compliance": [
        "esos", "secr", "carbon reporting", "carbon footprint", "net zero",
        "decarbonisation", "decarbonization", "energy audit", "energy audits",
        "display energy certificate", "energy performance certificate", "epc", "dec",
        "iso 50001", "heat decarbonisation plan", "carbon management plan",
        "scope 3", "greenhouse gas",
    ],
}

# Notices containing these are dropped.
EXCLUDE_KEYWORDS = ["fuel card", "fuel cards", "petrol", "diesel", "bottled gas", "lpg"]

# CPV code prefixes and how much a match adds to the score.
# Broad codes score low so they only pass when keywords also match.
CPV_WEIGHTS = {
    "71314": 5,     # Energy and related services (incl. energy mgmt, efficiency consultancy)
    "09310": 3,     # Electricity
    "09123": 3,     # Natural gas
    "0933": 4,      # Solar energy / panels / installation
    "79418": 2,     # Procurement consultancy
    "093": 1,       # Electricity, heating, solar and nuclear energy (broad)
    "65": 1,        # Public utilities (broad)
    "79411": 1,     # General management consultancy (broad)
    "90713": 1,     # Environmental issues consultancy (broad)
    "71313": 1,     # Environmental engineering consultancy (broad)
}

TITLE_HIT = 4        # per keyword found in the title (capped)
DESC_HIT = 1         # per keyword found in the description (capped)
DEFAULT_MIN_SCORE = 4

FTS_URL = "https://www.find-tender.service.gov.uk/api/1.0/ocdsReleasePackages"
CF_URL = "https://www.contractsfinder.service.gov.uk/Published/Notices/OCDS/Search"
FTS_NOTICE = "https://www.find-tender.service.gov.uk/Notice/{}"
CF_NOTICE = "https://www.contractsfinder.service.gov.uk/Notice/{}"

HEADERS = {"Accept": "application/json", "User-Agent": "EPL-TenderFinder/1.0"}
MAX_PAGES = 600
SLEEP_BETWEEN = 0.4  # be polite to the APIs

# --------------------------------------------------------------------------
# Fetching
# --------------------------------------------------------------------------

def _get(session, url, params=None):
    """GET with retry/backoff on rate limits and transient errors."""
    for attempt in range(6):
        try:
            r = session.get(url, params=params, headers=HEADERS, timeout=60)
        except requests.RequestException as e:
            wait = 2 ** attempt
            print(f"  network error ({e}); retrying in {wait}s", file=sys.stderr)
            time.sleep(wait)
            continue
        if r.status_code == 429 or r.status_code >= 500:
            wait = int(r.headers.get("Retry-After", 2 ** attempt))
            print(f"  HTTP {r.status_code}; retrying in {wait}s", file=sys.stderr)
            time.sleep(wait)
            continue
        if r.status_code == 404:
            return None
        r.raise_for_status()
        return r.json()
    raise RuntimeError(f"Gave up on {url}")


def fetch_releases(source, date_from, date_to, stages):
    """Yield (source, release) for every release in the window, following pagination."""
    session = requests.Session()
    fmt = "%Y-%m-%dT%H:%M:%S"
    if source == "fts":
        url = FTS_URL
        params = {"updatedFrom": date_from.strftime(fmt), "updatedTo": date_to.strftime(fmt),
                  "stages": stages, "limit": 100}
    else:
        url = CF_URL
        params = {"publishedFrom": date_from.strftime(fmt), "publishedTo": date_to.strftime(fmt),
                  "stages": stages, "limit": 100}

    pages = 0
    while url and pages < MAX_PAGES:
        data = _get(session, url, params)
        pages += 1
        if not data:
            break
        releases = data.get("releases", [])
        for rel in releases:
            yield source, rel
        print(f"  {source.upper()} [{stages}] page {pages}: {len(releases)} notices")

        nxt = (data.get("links") or {}).get("next")
        if nxt and releases:
            url, params = nxt, None          # next link already carries the cursor
        elif data.get("nextCursor") and releases:
            params = dict(params or {}, cursor=data["nextCursor"])
        else:
            break
        time.sleep(SLEEP_BETWEEN)

# --------------------------------------------------------------------------
# Parsing helpers
# --------------------------------------------------------------------------

def parse_dt(s):
    if not s:
        return None
    try:
        dt = datetime.fromisoformat(str(s).replace("Z", "+00:00"))
        if dt.tzinfo:
            dt = dt.astimezone(timezone.utc).replace(tzinfo=None)
        return dt
    except ValueError:
        return None


def clean(text, limit=None):
    if not text:
        return ""
    text = re.sub(r"\s+", " ", str(text)).strip()
    if limit and len(text) > limit:
        text = text[: limit - 1] + "…"
    return text


_kw_cache = {}

def kw_regex(kw):
    if kw not in _kw_cache:
        _kw_cache[kw] = re.compile(r"(?<![a-z0-9])" + re.escape(kw.lower()) + r"(?![a-z0-9])")
    return _kw_cache[kw]


def find_keywords(text):
    """Return {category: [keywords]} found in text."""
    text = (text or "").lower()
    hits = {}
    for cat, kws in KEYWORD_GROUPS.items():
        found = [k for k in kws if kw_regex(k).search(text)]
        if found:
            hits[cat] = found
    return hits


def collect_cpvs(tender):
    codes = []
    def add(c):
        if c and c.get("scheme", "CPV").upper() == "CPV" and c.get("id"):
            codes.append(str(c["id"]).split("-")[0])
    add(tender.get("classification") or {})
    for item in tender.get("items") or []:
        add(item.get("classification") or {})
        for c in item.get("additionalClassifications") or []:
            add(c)
    for lot in tender.get("lots") or []:
        add(lot.get("classification") or {})
    return list(dict.fromkeys(codes))


def cpv_score(cpvs):
    best = 0
    for code in cpvs:
        for prefix, w in CPV_WEIGHTS.items():
            if code.startswith(prefix):
                best = max(best, w)
    return best


def buyer_party(rel):
    parties = rel.get("parties") or []
    buyer_id = (rel.get("buyer") or {}).get("id")
    for p in parties:
        if buyer_id and p.get("id") == buyer_id:
            return p
    for p in parties:
        if "buyer" in (p.get("roles") or []):
            return p
    b = rel.get("buyer") or {}
    return {"name": b.get("name"), "contactPoint": {}, "address": {}}


def other_contacts(rel, buyer):
    """Additional contact points (e.g. a procurement agent or central purchasing body)."""
    out = []
    for p in rel.get("parties") or []:
        if p is buyer:
            continue
        roles = p.get("roles") or []
        if any(r in roles for r in ("procuringEntity", "centralPurchasingBody", "reviewBody")) \
                and "reviewBody" not in roles:
            cp = p.get("contactPoint") or {}
            out.append(", ".join(x for x in [p.get("name"), cp.get("name"),
                                              cp.get("email"), cp.get("telephone")] if x))
    return "; ".join(out)


def criteria_text(block):
    parts = []
    for c in (block or {}).get("criteria") or []:
        label = c.get("name") or c.get("type") or ""
        weight = ""
        for n in c.get("numbers") or []:
            if n.get("number") is not None:
                weight = f" {n['number']}%" if n.get("weight") in (None, "percentageExact", "decimalExact") else f" {n['number']}"
                break
        desc = clean(c.get("description"), 120)
        parts.append(clean(f"{label}{weight}" + (f" ({desc})" if desc and desc != label else "")))
    return "; ".join(p for p in parts if p)


def notice_url(source, rel):
    if source == "fts":
        return FTS_NOTICE.format(rel.get("id", ""))
    ocid = rel.get("ocid", "")
    guid = ocid.split("ocds-b5fd17-")[-1] if "ocds-b5fd17-" in ocid else rel.get("id", "").rsplit("-", 1)[0]
    return CF_NOTICE.format(guid)


def stage_of(rel):
    tags = rel.get("tag") or []
    for t in ("planning", "tender", "award", "contract"):
        if any(str(x).startswith(t) for x in tags):
            return t
    return ",".join(tags)


def regions_of(tender, buyer):
    regs = []
    for item in tender.get("items") or []:
        for a in item.get("deliveryAddresses") or []:
            regs.append(a.get("region") or a.get("locality") or "")
        dl = item.get("deliveryLocation") or {}
        if dl.get("description"):
            regs.append(dl["description"])
    for lot in tender.get("lots") or []:
        for a in (lot.get("deliveryAddresses") or []):
            regs.append(a.get("region") or "")
    if not regs:
        regs.append((buyer.get("address") or {}).get("region") or "")
    return ", ".join(dict.fromkeys(r for r in regs if r))


def tender_value(tender):
    v = tender.get("value") or tender.get("maxValue") or {}
    if v.get("amount") is not None:
        return v.get("amount"), v.get("currency", "GBP")
    total, cur = 0, "GBP"
    for lot in tender.get("lots") or []:
        lv = lot.get("value") or {}
        if lv.get("amount"):
            total += lv["amount"]
            cur = lv.get("currency", cur)
    return (total or None), cur

# --------------------------------------------------------------------------
# Build a flat record from an OCDS release
# --------------------------------------------------------------------------

def to_record(source, rel):
    tender = rel.get("tender") or {}
    planning = rel.get("planning") or {}
    title = clean(tender.get("title") or rel.get("title") or planning.get("rationale"))
    desc = clean(tender.get("description") or rel.get("description"))
    text_all = f"{title} {desc}".lower()

    if any(kw_regex(x).search(text_all) for x in EXCLUDE_KEYWORDS):
        return None

    title_hits = find_keywords(title)
    desc_hits = find_keywords(desc)
    cpvs = collect_cpvs(tender)
    cs = cpv_score(cpvs)

    n_title = sum(len(v) for v in title_hits.values())
    n_desc = sum(len(v) for v in desc_hits.values())
    kw_score = min(n_title, 3) * TITLE_HIT + min(n_desc, 4) * DESC_HIT
    score = kw_score + cs

    cats = {}
    for h in (title_hits, desc_hits):
        for c, ks in h.items():
            cats.setdefault(c, set()).update(ks)
    if not cats and cs >= 4:
        cats = {"Energy (CPV match)": set()}
    category = max(cats, key=lambda c: len(cats[c])) if cats else ""
    matched = sorted({k for ks in cats.values() for k in ks})

    buyer = buyer_party(rel)
    cp = buyer.get("contactPoint") or {}
    addr = buyer.get("address") or {}
    address = ", ".join(x for x in [addr.get("streetAddress"), addr.get("locality"),
                                    addr.get("region"), addr.get("postalCode")] if x)

    value, currency = tender_value(tender)
    period = tender.get("tenderPeriod") or {}
    enquiry = tender.get("enquiryPeriod") or {}
    cperiod = tender.get("contractPeriod") or {}
    lots = tender.get("lots") or []
    if not cperiod and lots:
        cperiod = lots[0].get("contractPeriod") or {}

    award_crit = criteria_text(tender.get("awardCriteria")) or \
        " | ".join(filter(None, (criteria_text(l.get("awardCriteria")) for l in lots)))
    select_crit = criteria_text(tender.get("selectionCriteria")) or \
        " | ".join(filter(None, (criteria_text(l.get("selectionCriteria")) for l in lots)))

    docs = [d.get("url") for d in (tender.get("documents") or []) if d.get("url")]
    for lot in lots:
        docs += [d.get("url") for d in (lot.get("documents") or []) if d.get("url")]
    portal = tender.get("submissionMethodDetails") or ""
    if not portal:
        for d in tender.get("documents") or []:
            if d.get("documentType") in ("tenderNotice", "biddingDocuments") and d.get("url"):
                portal = d["url"]
                break

    sme = tender.get("suitability", {}).get("sme") if tender.get("suitability") else None
    if sme is None and lots:
        sme = any((l.get("suitability") or {}).get("sme") for l in lots)

    duration = cperiod.get("durationInDays")
    if not duration and cperiod.get("startDate") and cperiod.get("endDate"):
        s, e = parse_dt(cperiod["startDate"]), parse_dt(cperiod["endDate"])
        duration = (e - s).days if s and e else None

    # Award info (for incumbents sheet)
    awards = rel.get("awards") or []
    suppliers = sorted({s.get("name") for a in awards for s in (a.get("suppliers") or []) if s.get("name")})
    award_val = sum((a.get("value") or {}).get("amount") or 0 for a in awards) or None
    award_date = next((parse_dt(a.get("date")) for a in awards if a.get("date")), None)
    end_dates = [parse_dt((a.get("contractPeriod") or {}).get("endDate")) for a in awards]
    end_dates += [parse_dt((c.get("period") or {}).get("endDate")) for c in rel.get("contracts") or []]
    end_dates = [d for d in end_dates if d] or ([parse_dt(cperiod.get("endDate"))] if cperiod.get("endDate") else [])
    contract_end = max(end_dates) if end_dates else None

    return {
        "source": "Find a Tender" if source == "fts" else "Contracts Finder",
        "ocid": rel.get("ocid"),
        "stage": stage_of(rel),
        "score": score,
        "category": category,
        "matched": ", ".join(matched),
        "title": title,
        "buyer": clean(buyer.get("name")),
        "contact_name": clean(cp.get("name")),
        "email": clean(cp.get("email")),
        "phone": clean(cp.get("telephone")),
        "website": clean(cp.get("url") or (buyer.get("details") or {}).get("url")),
        "address": clean(address),
        "region": regions_of(tender, buyer),
        "other_contacts": other_contacts(rel, buyer),
        "value": value,
        "currency": currency,
        "published": parse_dt(rel.get("date")),
        "deadline": parse_dt(period.get("endDate")),
        "clarification_deadline": parse_dt(enquiry.get("endDate")),
        "contract_start": parse_dt(cperiod.get("startDate")),
        "duration_days": duration,
        "procedure": clean(tender.get("procurementMethodDetails") or tender.get("procurementMethod")),
        "framework": "Yes" if (tender.get("techniques") or {}).get("hasFrameworkAgreement") else "",
        "lots": len(lots) or "",
        "sme": "Yes" if sme else "",
        "award_criteria": clean(award_crit, 600),
        "selection_criteria": clean(select_crit, 600),
        "cpvs": ", ".join(cpvs),
        "portal": clean(portal),
        "documents": " ".join(dict.fromkeys(docs))[:2000],
        "description": clean(desc, 1500),
        "url": notice_url(source, rel),
        "suppliers": ", ".join(suppliers),
        "award_value": award_val,
        "award_date": award_date,
        "contract_end": contract_end,
    }


def collect(sources, stages, date_from, date_to, min_score, region=None):
    latest = {}
    for src in sources:
        for stage in stages:
            print(f"Fetching {src.upper()} {stage} notices {date_from:%d %b %Y} -> {date_to:%d %b %Y}")
            for source, rel in fetch_releases(src, date_from, date_to, stage):
                try:
                    rec = to_record(source, rel)
                except Exception as e:  # never let one odd notice kill the run
                    print(f"  skipped a notice ({e})", file=sys.stderr)
                    continue
                if not rec or rec["score"] < min_score:
                    continue
                if region:
                    blob = f"{rec['region']} {rec['address']}".lower()
                    if region.lower() not in blob:
                        continue
                key = rec["ocid"] or rec["url"]
                prev = latest.get(key)
                if not prev or (rec["published"] or datetime.min) > (prev["published"] or datetime.min):
                    latest[key] = rec

    # Cross-source de-dup: same buyer + title published on both portals
    seen, out = set(), []
    for rec in sorted(latest.values(), key=lambda r: r["source"] != "Find a Tender"):
        k = (rec["buyer"].lower(), re.sub(r"\W+", "", rec["title"].lower())[:80])
        if k in seen:
            continue
        seen.add(k)
        out.append(rec)
    return out

# --------------------------------------------------------------------------
# Excel output
# --------------------------------------------------------------------------

FONT = "Arial"
HEAD_FILL = PatternFill("solid", start_color="1F4E3D")
HEAD_FONT = Font(name=FONT, bold=True, color="FFFFFF")
BODY_FONT = Font(name=FONT, size=10)
LINK_FONT = Font(name=FONT, size=10, color="0563C1", underline="single")


def write_sheet(ws, columns, rows):
    """columns: list of (header, key_or_callable, width, number_format)."""
    for ci, (h, _, w, _) in enumerate(columns, 1):
        c = ws.cell(row=1, column=ci, value=h)
        c.font, c.fill = HEAD_FONT, HEAD_FILL
        c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[get_column_letter(ci)].width = w
    for ri, rec in enumerate(rows, 2):
        for ci, (h, key, _, fmt) in enumerate(columns, 1):
            val = key(rec, ri) if callable(key) else rec.get(key)
            c = ws.cell(row=ri, column=ci, value=val if val not in (None, "") else None)
            c.font = BODY_FONT
            c.alignment = Alignment(wrap_text=False, vertical="top")
            if fmt:
                c.number_format = fmt
            if isinstance(val, str) and val.startswith("http") and " " not in val:
                c.hyperlink, c.font = val, LINK_FONT
    ws.freeze_panes = "D2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{max(len(rows) + 1, 2)}"
    ws.row_dimensions[1].height = 32


def build_workbook(opps, awards, args, path):
    wb = Workbook()
    now = datetime.now()

    # Rank: open first, then score, then soonest deadline
    def rank(r):
        open_ = r["deadline"] is None or r["deadline"] >= now
        return (not open_, -r["score"], r["deadline"] or datetime.max)
    opps = sorted(opps, key=rank)

    # --- Opportunities
    ws = wb.active
    ws.title = "Opportunities"
    dl_col = 12  # column L = Deadline (keep in sync with the list below)
    cols = [
        ("Score", "score", 7, "0"),
        ("Category", "category", 20, None),
        ("Stage", "stage", 9, None),
        ("Title", "title", 50, None),
        ("Buyer", "buyer", 32, None),
        ("Contact name", "contact_name", 20, None),
        ("Email", "email", 30, None),
        ("Phone", "phone", 16, None),
        ("Region", "region", 16, None),
        ("Value", "value", 14, "£#,##0"),
        ("Published", "published", 12, "dd/mm/yyyy"),
        ("Deadline", "deadline", 16, "dd/mm/yyyy hh:mm"),
        ("Days left", lambda r, i: f'=IF(L{i}="","",INT(L{i}-NOW()))', 9, "0;[Red]-0"),
        ("Clarification deadline", "clarification_deadline", 16, "dd/mm/yyyy hh:mm"),
        ("Contract start", "contract_start", 12, "dd/mm/yyyy"),
        ("Duration (days)", "duration_days", 10, "0"),
        ("Procedure", "procedure", 18, None),
        ("Framework", "framework", 10, None),
        ("Lots", "lots", 6, None),
        ("SME suitable", "sme", 8, None),
        ("Award criteria", "award_criteria", 40, None),
        ("Selection criteria", "selection_criteria", 40, None),
        ("Matched keywords", "matched", 30, None),
        ("CPV codes", "cpvs", 18, None),
        ("Tender portal", "portal", 30, None),
        ("Notice link", "url", 30, None),
        ("Documents", "documents", 30, None),
        ("Buyer website", "website", 24, None),
        ("Buyer address", "address", 30, None),
        ("Other contacts", "other_contacts", 30, None),
        ("Description", "description", 80, None),
        ("Source", "source", 14, None),
    ]
    assert cols[dl_col - 1][0] == "Deadline"
    write_sheet(ws, cols, opps)
    if opps:
        rng = f"M2:M{len(opps) + 1}"
        ws.conditional_formatting.add(rng, CellIsRule(operator="between", formula=["0", "14"],
                                      fill=PatternFill("solid", start_color="FCE4D6")))
        ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["14"],
                                      fill=PatternFill("solid", start_color="E2EFDA")))

    # --- Buyer Contacts (de-duplicated)
    contacts = {}
    for r in opps + awards:
        if not r["buyer"]:
            continue
        k = (r["buyer"].lower(), (r["email"] or r["contact_name"]).lower())
        c = contacts.setdefault(k, {**r, "count": 0, "titles": []})
        c["count"] += 1
        if len(c["titles"]) < 3:
            c["titles"].append(r["title"])
        for f in ("contact_name", "email", "phone", "website", "address"):
            c[f] = c[f] or r[f]
    crow = sorted(contacts.values(), key=lambda c: (-c["count"], c["buyer"]))
    for c in crow:
        c["titles"] = " | ".join(c["titles"])
    write_sheet(wb.create_sheet("Buyer Contacts"), [
        ("Buyer", "buyer", 34, None),
        ("Contact name", "contact_name", 22, None),
        ("Email", "email", 32, None),
        ("Phone", "phone", 16, None),
        ("Website", "website", 28, None),
        ("Address", "address", 40, None),
        ("Region", "region", 16, None),
        ("Relevant notices", "count", 9, "0"),
        ("Example notices", "titles", 70, None),
    ], crow)

    # --- Incumbents / renewals
    if awards:
        awards = sorted(awards, key=lambda r: r["contract_end"] or datetime.max)
        write_sheet(wb.create_sheet("Incumbents"), [
            ("Category", "category", 20, None),
            ("Title", "title", 50, None),
            ("Buyer", "buyer", 32, None),
            ("Winning supplier(s)", "suppliers", 32, None),
            ("Award value", "award_value", 14, "£#,##0"),
            ("Award date", "award_date", 12, "dd/mm/yyyy"),
            ("Contract end", "contract_end", 12, "dd/mm/yyyy"),
            ("Months to end", lambda r, i: f'=IF(G{i}="","",ROUND((G{i}-TODAY())/30.4,0))', 9, "0;[Red]-0"),
            ("Contact name", "contact_name", 20, None),
            ("Email", "email", 30, None),
            ("Phone", "phone", 16, None),
            ("Region", "region", 16, None),
            ("Notice link", "url", 30, None),
            ("Source", "source", 14, None),
        ], awards)

    # --- About
    ab = wb.create_sheet("About")
    lines = [
        ("Generated", now.strftime("%d %b %Y %H:%M")),
        ("Window", f"Last {args.days} days (tenders/pipeline)" +
                   (f", last {args.award_days} days (awards)" if args.awards else "")),
        ("Sources", "Find a Tender Service + Contracts Finder (official OCDS APIs)"),
        ("Minimum score", args.min_score),
        ("Region filter", args.region or "None"),
        ("Opportunities", len(opps)),
        ("Score", "Keyword hits in title (x4, max 3) + description (x1, max 4) + best CPV weight"),
        ("Days left", "Green = more than 14 days, amber = 14 days or fewer, red = closed"),
        ("Stage", "'planning' = pipeline / early market engagement notice - tender not yet live"),
        ("Contacting buyers", "During a live tender, use the portal's clarification route only - "
                              "canvassing the buyer outside it can disqualify a bid."),
        ("Full ITT documents", "Usually held on the buyer's e-tendering portal (see 'Tender portal' / "
                               "'Notice link'); registration is often required."),
        ("Tuning", "Edit KEYWORD_GROUPS, EXCLUDE_KEYWORDS and CPV_WEIGHTS at the top of the script."),
    ]
    for i, (k, v) in enumerate(lines, 1):
        ab.cell(row=i, column=1, value=k).font = Font(name=FONT, bold=True)
        ab.cell(row=i, column=2, value=v).font = BODY_FONT
    ab.column_dimensions["A"].width = 22
    ab.column_dimensions["B"].width = 110

    wb.save(path)

# --------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description="Find UK public energy tenders for EPL")
    ap.add_argument("--days", type=int, default=30, help="look-back window for tenders (default 30)")
    ap.add_argument("--min-score", type=int, default=DEFAULT_MIN_SCORE)
    ap.add_argument("--open-only", action="store_true", help="drop tenders whose deadline has passed")
    ap.add_argument("--no-planning", action="store_true", help="skip pipeline/early engagement notices")
    ap.add_argument("--region", help='filter on region text or NUTS code, e.g. "North East" or UKC')
    ap.add_argument("--sources", default="fts,cf", help="fts, cf or fts,cf (default both)")
    ap.add_argument("--awards", action="store_true", help="also pull award notices for incumbent intel")
    ap.add_argument("--award-days", type=int, default=730, help="look-back for awards (default 730)")
    ap.add_argument("--out", default=None, help="output .xlsx path")
    ap.add_argument("--site-dir", default=None, help="also publish a web page + data into this folder (e.g. docs)")
    ap.add_argument("--email", action="store_true", help="email new matches (needs SMTP_* environment variables)")
    args = ap.parse_args()

    sources = [s.strip() for s in args.sources.split(",") if s.strip() in ("fts", "cf")]
    now = datetime.now(timezone.utc).replace(tzinfo=None)
    stages = ["tender"] + ([] if args.no_planning else ["planning"])

    opps = collect(sources, stages, now - timedelta(days=args.days), now, args.min_score, args.region)
    if args.open_only:
        opps = [r for r in opps if r["deadline"] is None or r["deadline"] >= datetime.now()]

    awards = []
    if args.awards:
        # Pull awards in 90-day chunks so long windows stay manageable
        start = now - timedelta(days=args.award_days)
        while start < now:
            end = min(start + timedelta(days=90), now)
            awards += collect(sources, ["award"], start, end, args.min_score, args.region)
            start = end
        seen, uniq = set(), []
        for a in awards:  # an award updated in two windows shows up twice
            k = a["ocid"] or a["url"]
            if k not in seen:
                seen.add(k)
                uniq.append(a)
        awards = uniq

    if args.site_dir:
        import site_builder
        if not args.awards:  # keep last known renewals between weekly award scans
            awards = site_builder.load_previous_awards(args.site_dir)
        new_items = site_builder.mark_new(args.site_dir, opps)
        out = args.out or f"{args.site_dir}/EPL_Tenders_latest.xlsx"
        build_workbook(opps, awards, args, out)
        site_builder.write_site(args.site_dir, opps, awards, new_items)
        if args.email:
            site_builder.send_alert(new_items)
    else:
        out = args.out or f"EPL_Tenders_{datetime.now():%Y-%m-%d}.xlsx"
        build_workbook(opps, awards, args, out)
    print(f"\nDone: {len(opps)} opportunities" + (f", {len(awards)} awards" if args.awards else "")
          + f" -> {out}")


if __name__ == "__main__":
    main()

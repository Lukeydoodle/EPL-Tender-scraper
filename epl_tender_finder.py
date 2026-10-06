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
# Each group is one EPL service line. Add or remove phrases freely.
KEYWORD_GROUPS = {
    "Energy Procurement / Brokerage": [
        "energy broker", "energy brokerage", "energy brokers", "utility broker", "utilities broker",
        "utility brokerage", "utilities brokerage", "energy procurement", "utilities procurement",
        "utility procurement", "procurement of energy", "procurement of utilities",
        "energy purchasing", "flexible procurement", "flexible energy", "energy basket",
        "energy consultancy", "energy consultant", "energy consultants", "utility consultancy",
        "utilities consultancy", "utilities consultant", "energy advisory", "energy advisor",
        "energy and utilities", "energy services", "utilities management", "utility management",
        "energy contract management", "risk management energy", "electricity and gas",
        "gas and electricity", "energy supply contracts", "tariff review",
    ],
    "Bill Validation / Bureau": [
        "bill validation", "invoice validation", "utility bill", "utility bills",
        "energy invoice", "energy invoices", "billing validation", "bill auditing",
        "utility audit", "utilities audit", "cost recovery", "overcharge recovery",
        "bureau service", "bureau services", "data bureau", "utilities bureau",
        "energy bureau", "invoice management",
    ],
    "Energy Supply": [
        "supply of electricity", "supply of gas", "electricity supply", "gas supply",
        "natural gas", "half hourly", "non half hourly", "unmetered supply",
        "utilities supply", "renewable electricity", "green electricity", "rego",
    ],
    "Metering / Energy Management": [
        "energy management", "energy management system", "energy monitoring",
        "monitoring and targeting", "sub-metering", "sub metering", "submetering", "sub-meters",
        "smart meter", "smart meters", "smart metering", "automatic meter reading", "amr",
        "meter data", "metering services", "half hourly data", "energy dashboard",
        "energy data", "energy analytics", "meter operator", "data collection",
    ],
    "Solar / BESS / EV": [
        "solar", "solar pv", "photovoltaic", "photovoltaics", "solar panels", "solar farm",
        "solar carport", "rooftop solar", "battery storage", "bess", "battery energy storage",
        "energy storage", "ev charging", "electric vehicle charging", "ev chargers",
        "ev infrastructure", "chargepoint", "chargepoints", "charge point", "charge points",
        "charging infrastructure", "electric vehicle infrastructure", "private wire",
        "renewable energy", "renewables", "on-site generation", "onsite generation",
    ],
    "CPPA / PPA": [
        "power purchase agreement", "power purchase agreements", "ppa", "cppa",
        "corporate ppa", "sleeved ppa", "virtual ppa", "offsite ppa",
    ],
    "Carbon / Compliance": [
        "esos", "energy savings opportunity scheme", "secr",
        "streamlined energy and carbon reporting", "carbon reporting", "carbon accounting",
        "carbon footprint", "carbon footprinting", "carbon management", "carbon management plan",
        "greenhouse gas", "ghg", "scope 1", "scope 2", "scope 3", "net zero", "net-zero",
        "net zero strategy", "net zero roadmap", "decarbonisation", "decarbonization",
        "decarbonisation plan", "heat decarbonisation plan", "carbon reduction",
        "sustainability reporting", "energy audit", "energy audits", "energy survey",
        "energy surveys", "energy assessment", "energy assessments",
        "display energy certificate", "display energy certificates",
        "energy performance certificate", "energy performance certificates", "epc", "epcs", "dec", "decs",
        "mees", "minimum energy efficiency standards", "iso 50001", "enms",
        "energy efficiency", "energy saving", "energy savings", "psds",
        "public sector decarbonisation scheme",
    ],
}

# Notices containing these are dropped (building works, not consultancy/brokerage).
EXCLUDE_KEYWORDS = [
    "fuel card", "fuel cards", "petrol", "diesel", "bottled gas", "lpg", "heating oil",
    "insulation", "cavity wall", "external wall insulation", "window replacement",
    "boiler replacement", "boiler servicing", "roofing", "kitchen replacement",
    "bathroom replacement", "gas servicing", "gas safety", "street lighting maintenance",
    "solar eclipse",
]

# CPV code prefixes and how much a match adds to the score.
# Broad codes score low so they only pass when keywords also match.
CPV_WEIGHTS = {
    "71314": 5,     # Energy and related services (energy management, efficiency consultancy)
    "0933": 4,      # Solar energy / panels / installation
    "09310": 3,     # Electricity
    "09123": 3,     # Natural gas
    "38554": 3,     # Electricity meters
    "31681500": 3,  # Rechargers (EV charging)
    "31158": 2,     # Chargers
    "31440": 2,     # Batteries (storage)
    "38550": 2,     # Meters
    "79418": 2,     # Procurement consultancy
    "093": 1,       # Electricity, heating, solar and nuclear energy (broad)
    "65": 1,        # Public utilities (broad)
    "79411": 1,     # General management consultancy (broad)
    "79212": 1,     # Auditing services (broad)
    "90713": 1,     # Environmental issues consultancy (broad)
    "90714": 1,     # Environmental auditing (broad)
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
    guid = ocid.split("ocds-b5fd17-")[-1] if "ocds-b5fd17-" in ocid else
 if __name__ == "__main__":
    main()

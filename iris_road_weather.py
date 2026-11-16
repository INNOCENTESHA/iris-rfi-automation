#!/usr/bin/env python3
"""
IRIS Road Weather Advisory
==========================
Daily construction road-weather advisory for the IRIS-LUPTAN JV road projects
on Unguja and Pemba (Zanzibar).

Pipeline (per island):
  1. Read the RFI_DATA tabs of both project spreadsheets (Google Sheets API)
  2. Scope to site works inspected in the last 3 days
  3. Resolve the in-scope road codes to coordinates from the master road list,
     preferring survey-derived centre coordinates where available
  4. Fetch the weather forecast for those roads only
  5. Compute STATUS and icons deterministically from fixed thresholds
  6. Render the complete HTML report (fixed sections, colours and fonts)
  7. Ask Claude for the narrative parts only (executive summary + advisories)
  8. Email the report

Rendering is deterministic: Python builds every table and status badge, and the
model writes prose only, so the report cannot be truncated or restyled.

Config files (weather_assets/, next to this script):
  <island>_roads.json    road names + coordinates
  spec_clauses.json      SSRW 2000 / PMDM 1999 clauses cited in the advisories
  recipients.json        {"unguja":{"to":[...]},"pemba":{"to":[...]}}  (optional;
                         WEATHER_RECIPIENTS_JSON takes precedence)

Usage:
  python iris_road_weather.py --island unguja --dry-run
  python iris_road_weather.py --island both

Environment:
  ANTHROPIC_API_KEY        narrative (omit -> built-in narrative)
  ANTHROPIC_MODEL          default claude-sonnet-5-5
  GOOGLE_TOKEN_PATH        OAuth token file (default: google_token.json here)
  GMAIL_USER               sending account (required to send)
  GMAIL_APP_PASSWORD       app password for that account (required to send)
  WEATHER_RECIPIENTS_JSON  recipients as JSON (preferred over recipients.json)
  WEATHER_PROVIDER         openmeteo (default) | msn
  IRIS_BASE_DIR            base folder (default: this script's folder)
"""

import os
import re
import csv
import ssl
import sys
import json
import time
import html
import smtplib
import argparse
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta, timezone
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from concurrent.futures import ThreadPoolExecutor

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# ── Config ────────────────────────────────────────────────────────────────────
BASE_DIR = os.environ.get("IRIS_BASE_DIR", os.path.dirname(os.path.abspath(__file__)))
ASSETS = os.path.join(BASE_DIR, "weather_assets")
RECIPIENTS_FILE = os.path.join(ASSETS, "recipients.json")
TOKEN_FILE = os.environ.get("GOOGLE_TOKEN_PATH", os.path.join(BASE_DIR, "google_token.json"))
SPEC_FILES = {
    "SSRW 2000": os.path.join(BASE_DIR, "SSRW2000.csv"),
    "PMDM 1999": os.path.join(BASE_DIR, "PMDM1999.csv"),
}

SS_163 = "1IjN4W79VEZrAjevG5zZVKyJPuCdGbm_xAbcttTNjfOA"
SS_655 = "1WV58zo-jYgqOTBGOtMAjjlHC3WKdQ83s-XdRFlxZLLU"

EAT = timezone(timedelta(hours=3))          # Africa/Dar_es_Salaam, no DST
SLOTS = [6, 9, 12, 15, 18, 21]              # fixed 3-hourly EAT slots
NIGHT_SLOTS = {18, 21}

ISLANDS = {
    "unguja": {
        "label": "UNGUJA (ZANZIBAR)",
        "roads": os.path.join(ASSETS, "unguja_roads.json"),
        "sheets": [(SS_163, "163.22 UNGUJA RFI_DATA"), (SS_655, "65.5 UNGUJA RFI_DATA")],
        "subject": "{date} UNGUJA _65.5 & 163.22 KM ROADS WEATHER ADVISORY",
        "next": "05:45",
    },
    "pemba": {
        "label": "PEMBA",
        "roads": os.path.join(ASSETS, "pemba_roads.json"),
        "sheets": [(SS_163, "163.22 PEMBA RFI_DATA"), (SS_655, "65.5 PEMBA RFI_DATA")],
        "subject": "{date} PEMBA _65.5 & 163.22 KM ROADS WEATHER ADVISORY",
        "next": "05:30",
    },
}

# Exact palette from the approved format spec.
CLR = {
    "page": "#f4f6f9", "card": "#ffffff", "body": "#1f2937",
    "navy": "#1a2e4a", "white": "#ffffff",
    "row_a": "#f8fafc", "row_b": "#ffffff", "border": "#e2e8f0",
}
STATUS_CLR = {
    "GOOD":     ("#dcfce7", "#166534"),
    "MODERATE": ("#fef9c3", "#854d0e"),
    "SEVERE":   ("#ffedd5", "#9a3412"),
    "EXTREME":  ("#fee2e2", "#991b1b"),
    "BAD":      ("#4c0519", "#ffffff"),
}
STATUS_ORDER = ["BAD", "EXTREME", "SEVERE", "MODERATE", "GOOD"]
STATUS_DESC = {
    "GOOD": "Dry/clear, wind &lt;20 km/h, precip &lt;20%. All works proceed: earthworks, "
            "compaction, bitumen/asphalt, concrete pours and curing. Site safe and fully accessible.",
    "MODERATE": "Light/intermittent showers, wind 20–35 km/h, precip 20–50%. Most works continue "
                "with care; protect fresh concrete/bitumen and watch subgrade moisture.",
    "SEVERE": "Moderate rain, wind 35–50 km/h, precip 50–70%. Suspend bitumen/asphalt and concrete "
              "pours; earthworks limited by wet subgrade; prioritise drainage.",
    "EXTREME": "Heavy rain/storms, wind 50–65 km/h, precip 70–90%. Halt earthworks, paving and "
               "pours; flooding/collapse risk. Hazardous access; lightning risk to crews.",
    "BAD": "Violent storms/flooding, wind &gt;65 km/h, precip &gt;90%. Stop all site works; clear "
           "low-lying excavations; protect materials and plant. Site unsafe/impassable.",
}
FONT = "Segoe UI, Tahoma, Geneva, Verdana, sans-serif"

# WMO weather code -> (condition text, day icon, night icon)
WMO = {
    0: ("Clear", "☀️", "🌙"), 1: ("Mostly sunny", "🌤️", "🌙"),
    2: ("Partly cloudy", "⛅", "🌙"), 3: ("Overcast", "☁️", "☁️"),
    45: ("Fog", "🌫️", "🌫️"), 48: ("Freezing fog", "🌫️", "🌫️"),
    51: ("Light drizzle", "🌦️", "🌦️"), 53: ("Drizzle", "🌦️", "🌦️"),
    55: ("Heavy drizzle", "🌦️", "🌦️"),
    61: ("Light rain", "🌦️", "🌦️"), 63: ("Rain", "🌧️", "🌧️"),
    65: ("Heavy rain", "🌧️", "🌧️"),
    66: ("Freezing rain", "🌧️", "🌧️"), 67: ("Freezing rain", "🌧️", "🌧️"),
    71: ("Light snow", "❄️", "❄️"), 73: ("Snow", "❄️", "❄️"), 75: ("Heavy snow", "❄️", "❄️"),
    77: ("Snow grains", "❄️", "❄️"),
    80: ("Scattered showers", "🌦️", "🌦️"), 81: ("Rain showers", "🌧️", "🌧️"),
    82: ("Violent rain showers", "🌧️", "🌧️"),
    85: ("Snow showers", "❄️", "❄️"), 86: ("Snow showers", "❄️", "❄️"),
    95: ("Thunderstorm", "⛈️", "⛈️"),
    96: ("Thunderstorm with hail", "⛈️", "⛈️"),
    99: ("Severe thunderstorm", "⛈️", "⛈️"),
}


def log(m=""):
    print(m, flush=True)


# ── Status & icons (documented thresholds — no model involvement) ─────────────
def status_for(precip_pct, wind_kmh, code=None):
    """Worst of precipitation / wind / storm severity."""
    p, w = float(precip_pct or 0), float(wind_kmh or 0)
    if p > 90 or w > 65:
        s = "BAD"
    elif p >= 70 or w >= 50:
        s = "EXTREME"
    elif p >= 50 or w >= 35:
        s = "SEVERE"
    elif p >= 20 or w >= 20:
        s = "MODERATE"
    else:
        s = "GOOD"
    # A thunderstorm is never merely MODERATE/GOOD.
    if code in (95, 96, 99) and STATUS_ORDER.index(s) > STATUS_ORDER.index("SEVERE"):
        s = "SEVERE"
    return s


def cond_icon(code, hour):
    """Condition text + icon, day/night aware (no sun icons at 18:00 / 21:00)."""
    text, day, night = WMO.get(int(code or 0), ("Cloudy", "☁️", "☁️"))
    icon = night if hour in NIGHT_SLOTS else day
    if hour in NIGHT_SLOTS and text in ("Clear", "Mostly sunny", "Partly cloudy"):
        text = "Clear night" if text in ("Clear", "Mostly sunny") else "Partly cloudy"
    return text, icon


def worst(statuses):
    for s in STATUS_ORDER:
        if s in statuses:
            return s
    return "GOOD"


def badge(status):
    bg, fg = STATUS_CLR[status]
    return (f'<span style="display:inline-block;padding:3px 12px;border-radius:999px;'
            f'font-weight:700;font-size:12px;background:{bg};color:{fg}">{status}</span>')


# ── Google Sheets ─────────────────────────────────────────────────────────────
def sheets_service():
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build
    creds = Credentials.from_authorized_user_file(TOKEN_FILE)
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
        with open(TOKEN_FILE, "w", encoding="utf-8") as f:
            f.write(creds.to_json())
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


def read_tab(svc, ssid, tab):
    try:
        vals = svc.spreadsheets().values().get(
            spreadsheetId=ssid, range=f"'{tab}'!A:Z").execute().get("values", [])
    except Exception as e:
        log(f"    ! could not read '{tab}': {e}")
        return []
    if not vals:
        return []
    # Header row = the row containing 'RFI No.' (tables sit below a title block).
    hi = 0
    for i, row in enumerate(vals[:40]):
        if "rfi no" in " ".join(str(c) for c in row).lower():
            hi = i
            break
    hdr = [str(c).strip() for c in vals[hi]]
    out = []
    for row in vals[hi + 1:]:
        if not any(str(c).strip() for c in row):
            continue
        out.append({hdr[i]: (row[i] if i < len(row) else "") for i in range(len(hdr))})
    return out


# ── RFI scoping ───────────────────────────────────────────────────────────────
MONTHS = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"])}


def parse_date(s):
    if not s:
        return None
    s = str(s).strip()
    m = re.match(r"^(\d{1,2})[/\-](\d{1,2})[/\-](\d{2,4})$", s)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        y += 2000 if y < 100 else 0
        try:
            return datetime(y, mo, d, tzinfo=EAT)
        except ValueError:
            return None
    m2 = re.match(r"^(\d{1,2})[/\-]([A-Za-z]{3})[/\-](\d{2,4})$", s)
    if m2:
        mo = MONTHS.get(m2.group(2).lower())
        if mo is not None:
            y = int(m2.group(3))
            y += 2000 if y < 100 else 0
            try:
                return datetime(y, mo + 1, int(m2.group(1)), tzinfo=EAT)
            except ValueError:
                return None
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d %b %Y", "%b %d, %Y"):
        try:
            return datetime.strptime(s[:11].strip(), fmt).replace(tzinfo=EAT)
        except Exception:
            pass
    return None


def scope_rfis(rows, days=3, now=None):
    now = now or datetime.now(EAT)
    cutoff = now - timedelta(days=days)
    horizon = now + timedelta(days=1)
    rec = []
    for r in rows:
        dstr = r.get("Inspection Date") or r.get("Date Received") or ""
        d = parse_date(dstr)
        if not d or d < cutoff or d > horizon:
            continue
        code = str(r.get("Road Code", "")).strip()
        name = str(r.get("ROAD NAME") or r.get("Road Name") or "").strip()
        desc = str(r.get("Work Description", "")).strip()
        if not code and not name and not desc:
            continue
        rec.append({
            "d": d, "code": code, "name": name,
            "act": str(r.get("Activity Name", "")).strip(), "desc": desc,
            "stn": str(r.get("Station Range", "")).strip(),
            "status": str(r.get("Status", "")).strip(),
            "rfi": str(r.get("RFI No.", "")).strip(), "dStr": dstr,
        })
    rec.sort(key=lambda x: x["d"], reverse=True)
    seen, uniq = set(), []
    for x in rec:
        k = f"{x['code']}|{x['desc']}|{x['stn']}"
        if k in seen:
            continue
        seen.add(k)
        uniq.append(x)
    lines = [f"- [{x['dStr']}] {x['code'] or 'N/A'} {x['name']} | {x['act'] or x['desc']} "
             f"({x['desc']}) | Ch {x['stn'] or '-'} | RFI {x['rfi']} "
             f"| RE: {x['status'] or 'OPEN'}" for x in uniq[:80]]
    brief = "\n".join(lines) if lines else "(No site-work RFIs in the last 3 days.)"
    codes = []
    for x in uniq:
        if x["code"] and x["code"] not in codes:
            codes.append(x["code"])
    return {"brief": brief, "count": len(uniq), "codes": codes, "items": uniq}


def roads_in_scope(island, codes):
    """Master list gives verbatim names + coordinates; every alignment counts.
    Survey centre coordinates are preferred when present."""
    master = json.load(open(ISLANDS[island]["roads"], encoding="utf-8"))
    for r in master:
        r["lat"] = r.get("lat_survey", r["lat"])
        r["lon"] = r.get("lon_survey", r["lon"])
        r["coord_src"] = "survey" if "lat_survey" in r else "master"
    if not codes:
        return []
    want = {c.upper() for c in codes}
    return [r for r in master if r["name"].split()[0].upper() in want]


# ── Weather providers ─────────────────────────────────────────────────────────
UA = {"User-Agent": "IRIS-RoadWeather/1.0 (construction advisory)"}


def _get(url, timeout=30):
    req = urllib.request.Request(url, headers=UA)
    with urllib.request.urlopen(req, timeout=timeout,
                               context=ssl.create_default_context()) as r:
        return r.read()


def fetch_openmeteo(lat, lon):
    """Genuine hourly series — temp, humidity, precip probability, wind, code."""
    q = urllib.parse.urlencode({
        "latitude": lat, "longitude": lon,
        "hourly": "temperature_2m,relative_humidity_2m,precipitation_probability,"
                  "wind_speed_10m,weather_code",
        "daily": "weather_code,temperature_2m_max,temperature_2m_min,"
                 "precipitation_probability_max,wind_speed_10m_max",
        "current": "temperature_2m,relative_humidity_2m,wind_speed_10m,weather_code",
        "timezone": "Africa/Nairobi", "forecast_days": 7, "wind_speed_unit": "kmh",
    })
    d = json.loads(_get("https://api.open-meteo.com/v1/forecast?" + q))
    H = d["hourly"]
    today = datetime.now(EAT).strftime("%Y-%m-%d")
    idx = {t: i for i, t in enumerate(H["time"])}
    hourly = []
    for h in SLOTS:
        i = idx.get(f"{today}T{h:02d}:00")
        if i is None:
            continue
        code = H["weather_code"][i]
        text, icon = cond_icon(code, h)
        precip = H["precipitation_probability"][i] or 0
        wind = H["wind_speed_10m"][i] or 0
        hourly.append({"hour": h, "temp": H["temperature_2m"][i], "text": text, "icon": icon,
                       "humidity": H["relative_humidity_2m"][i], "wind": round(wind),
                       "precip": round(precip), "code": code,
                       "status": status_for(precip, wind, code)})
    cur = d.get("current", {})
    ctext, cicon = cond_icon(cur.get("weather_code"), datetime.now(EAT).hour)
    D = d["daily"]
    daily = []
    for i, day in enumerate(D["time"][:7]):
        code = D["weather_code"][i]
        text, icon = cond_icon(code, 12)
        p = D["precipitation_probability_max"][i] or 0
        w = D["wind_speed_10m_max"][i] or 0
        dt = datetime.strptime(day, "%Y-%m-%d")
        daily.append({"day": dt.strftime("%a"), "date": dt.strftime("%d-%m-%Y"),
                      "text": text, "icon": icon, "high": D["temperature_2m_max"][i],
                      "low": D["temperature_2m_min"][i], "precip": round(p),
                      "wind": round(w), "status": status_for(p, w, code)})
    cw = cur.get("wind_speed_10m") or 0
    cp = hourly[0]["precip"] if hourly else 0
    return {"hourly": hourly, "daily": daily, "modelled": False,
            "current": {"temp": cur.get("temperature_2m"), "text": ctext, "icon": cicon,
                        "humidity": cur.get("relative_humidity_2m"), "wind": round(cw),
                        "status": status_for(cp, cw, cur.get("weather_code"))}}


MSN_SKY = {
    "sunny": 0, "clear": 0, "mostlysunny": 1, "partlysunny": 2, "partlycloudy": 2,
    "mostlycloudy": 3, "cloudy": 3, "fog": 45, "lightrain": 61, "rain": 63,
    "rainshowers": 81, "showers": 81, "tstorms": 95, "thunderstorms": 95,
}


def fetch_msn(lat, lon):
    """MSN Weather provider.

    IMPORTANT: this endpoint returns CURRENT conditions and a DAILY forecast
    only — it has no hourly series. The six hourly slots are therefore MODELLED
    from the day's high/low with a diurnal curve and labelled as such.
    """
    q = urllib.parse.urlencode({"src": "outlook", "weadegreetype": "C", "culture": "en-US",
                                "weasearchstr": f"{lat},{lon}", "count": "1"})
    root = ET.fromstring(_get("https://weather.service.msn.com/find.aspx?" + q))
    cur = root.find(".//current")
    fcs = root.findall(".//forecast")
    if cur is None:
        raise RuntimeError("MSN: no <current> element")

    def code_of(sky):
        return MSN_SKY.get(re.sub(r"[^a-z]", "", (sky or "").lower()), 3)

    def num(x, d=0.0):
        m = re.sub(r"[^\d.\-]", "", str(x or ""))
        try:
            return float(m)
        except ValueError:
            return d

    ccode = code_of(cur.get("skytext"))
    ctemp = num(cur.get("temperature"))
    cwind = num(cur.get("winddisplay"))
    chum = num(cur.get("humidity"))
    today = fcs[0] if fcs else None
    hi = num(today.get("high"), ctemp + 3) if today is not None else ctemp + 3
    lo = num(today.get("low"), ctemp - 3) if today is not None else ctemp - 3
    pr = num(today.get("precip")) if today is not None else 0
    dcode = code_of(today.get("skytextday")) if today is not None else ccode
    shape = {6: 0.0, 9: 0.55, 12: 0.9, 15: 1.0, 18: 0.6, 21: 0.3}   # cool 06:00, peak 15:00
    hourly = []
    for h in SLOTS:
        code = dcode if (h not in NIGHT_SLOTS or dcode >= 45) else 2
        text, icon = cond_icon(code, h)
        hourly.append({"hour": h, "temp": round(lo + (hi - lo) * shape[h], 1),
                       "text": text + " (modelled)", "icon": icon, "humidity": round(chum),
                       "wind": round(cwind), "precip": round(pr), "code": code,
                       "status": status_for(pr, cwind, code)})
    daily = []
    for f in fcs[:7]:
        code = code_of(f.get("skytextday"))
        text, icon = cond_icon(code, 12)
        p = num(f.get("precip"))
        try:
            dt = datetime.strptime(f.get("date"), "%Y-%m-%d")
            day, date = dt.strftime("%a"), dt.strftime("%d-%m-%Y")
        except Exception:
            day, date = f.get("day", ""), f.get("date", "")
        daily.append({"day": day, "date": date, "text": text, "icon": icon,
                      "high": f.get("high"), "low": f.get("low"), "precip": round(p),
                      "wind": round(cwind), "status": status_for(p, cwind, code)})
    ctext, cicon = cond_icon(ccode, datetime.now(EAT).hour)
    return {"hourly": hourly, "daily": daily, "modelled": True,
            "current": {"temp": ctemp, "text": ctext, "icon": cicon, "humidity": round(chum),
                        "wind": round(cwind), "status": status_for(pr, cwind, ccode)}}


def fetch_weather(road, provider):
    fn = fetch_msn if provider == "msn" else fetch_openmeteo
    last = None
    for attempt in range(3):
        try:
            w = fn(road["lat"], road["lon"])
            if w["hourly"]:
                return road, w, None
            last = "no hourly slots returned"
        except Exception as e:
            last = f"{type(e).__name__}: {e}"
        time.sleep(1.2 * (attempt + 1))
    return road, None, last


# ── Spec clause lookup ────────────────────────────────────────────────────────
CITED = {
    "SSRW 2000": ["3606", "3504", "3506", "3507", "3704", "3803", "4104", "4205"],
    "PMDM 1999": ["2.1", "2.2", "6.3.2"],
}


def load_clauses(limit_chars=700):
    """Prefer the pre-extracted clause bundle (small, CI-friendly); fall back to
    scanning the full SSRW/PMDM CSVs when it is absent."""
    bundle = os.path.join(ASSETS, "spec_clauses.json")
    if os.path.exists(bundle):
        try:
            got = json.load(open(bundle, encoding="utf-8"))
            if got:
                return got
        except Exception as e:
            log(f"  ! spec bundle unreadable ({e}) — falling back to CSVs")
    out = []
    for manual, path in SPEC_FILES.items():
        if not os.path.exists(path):
            log(f"  ! spec file missing: {path}")
            continue
        want = set(CITED.get(manual, []))
        found = {}
        with open(path, encoding="utf-8", errors="replace") as fh:
            for row in csv.DictReader(fh):
                cl = (row.get("clause") or "").strip()
                if cl in want and cl not in found:
                    found[cl] = {"manual": manual, "clause": cl,
                                 "title": (row.get("title") or "").strip(),
                                 "text": re.sub(r"\s+", " ", row.get("text") or "")[:limit_chars]}
                if len(found) == len(want):
                    break
        out.extend(found[c] for c in CITED[manual] if c in found)
    return out


# ── Claude narrative (Sections 2 & 4 prose only) ──────────────────────────────
NARR_SYS = (
    "You are the IRIS-LUPTAN JV Road Weather Intelligence analyst. You write ONLY the narrative "
    "prose of a construction road-weather advisory; all tables and styling are rendered "
    "separately, so emit no HTML beyond <p>, <ul>, <li>, <b>.\n"
    "All construction advice MUST follow Tanzania's Pavement and Materials Design Manual 1999 "
    "(PMDM 1999), adopted by Zanzibar, and cite the SSRW 2000 clause numbers supplied in the "
    "facts. Zanzibar lies in the PMDM WET zone (Table 2.1).\n"
    "Return STRICT JSON, no code fences, exactly these keys:\n"
    '{"executive_summary":"<p>...</p>","works_affected":"<ul>...</ul>",'
    '"site_safety":"<ul>...</ul>"}\n'
    "executive_summary: ONE short paragraph on dominant conditions, temperature range and "
    "highest-risk roads.\n"
    "works_affected: cover earthworks & compaction, bitumen/asphalt laying, concrete pours & "
    "curing, drainage/structures, material haulage & stockpiles — each justified against "
    "PMDM 1999 with the relevant SSRW 2000 clause in brackets.\n"
    "site_safety: crew safety (lightning, heat), plant & equipment, excavation/slope stability, "
    "flooding of works, haul-road condition."
)


def ai_narrative(island, rows, scope, clauses):
    key = os.environ.get("ANTHROPIC_API_KEY", "").strip()
    if not key:
        return None
    tally = {}
    for r in rows:
        tally[r["status"]] = tally.get(r["status"], 0) + 1
    temps = [h["temp"] for r in rows for h in r["w"]["hourly"] if h["temp"] is not None]
    facts = {
        "island": ISLANDS[island]["label"],
        "road_count": len(rows),
        "status_tally": tally,
        "temp_range_c": [min(temps), max(temps)] if temps else None,
        "worst_roads": [{"road": r["road"]["name"], "status": r["status"],
                         "max_precip_pct": max((h["precip"] or 0) for h in r["w"]["hourly"]),
                         "max_wind_kmh": max((h["wind"] or 0) for h in r["w"]["hourly"])}
                        for r in rows[:8]],
        "rfi_count": scope["count"],
        "rfi_brief": scope["brief"][:4000],
        "spec_clauses": clauses,
    }
    body = json.dumps({
        "model": os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-5-5"),
        "max_tokens": 3000, "system": NARR_SYS,
        "messages": [{"role": "user", "content": "FACTS (JSON):\n"
                      + json.dumps(facts, default=str)}],
    }).encode()
    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages", data=body,
        headers={"x-api-key": key, "anthropic-version": "2023-06-01",
                 "content-type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=180) as r:
            d = json.loads(r.read())
        txt = "".join(b.get("text", "") for b in d.get("content", []))
        txt = re.sub(r"^```(?:json)?|```$", "", txt.strip(), flags=re.M).strip()
        got = json.loads(txt)
        if all(k in got for k in ("executive_summary", "works_affected", "site_safety")):
            return got
        log("    ! narrative JSON missing keys — using built-in text")
    except Exception as e:
        log(f"    ! AI narrative unavailable ({type(e).__name__}: {e}) — using built-in text")
    return None


def fallback_narrative(rows):
    tally = {}
    for r in rows:
        tally[r["status"]] = tally.get(r["status"], 0) + 1
    temps = [h["temp"] for r in rows for h in r["w"]["hourly"] if h["temp"] is not None]
    rng = f"{min(temps):.0f}–{max(temps):.0f} °C" if temps else "n/a"
    dom = worst([r["status"] for r in rows]) if rows else "GOOD"
    hi = [r["road"]["name"] for r in rows if r["status"] in ("BAD", "EXTREME", "SEVERE")][:6]
    return {
        "executive_summary":
            f"<p>Across {len(rows)} road alignment(s) in scope the prevailing "
            f"construction-impact status is <b>{dom}</b>, with temperatures of {rng}. "
            f"Distribution: " + (", ".join(f"{k} {v}" for k, v in tally.items()) or "n/a")
            + (f". Highest-risk roads: {', '.join(html.escape(x) for x in hi)}."
               if hi else ". No high-risk roads today.") + "</p>",
        "works_affected":
            "<ul><li><b>Earthworks &amp; compaction</b> — place and compact at or near OMC; "
            "suspend where rain lifts material moisture above OMC, as specified density and CBR "
            "cannot then be achieved (PMDM 1999 Table 2.2; SSRW 2000 Cl. 3606, 3504, 3506).</li>"
            "<li><b>Bitumen/asphalt</b> — surface must be clean and dry; no prime or paving when "
            "wet, foggy or with rain imminent (SSRW 2000 Cl. 4104, 4205; Series 4300).</li>"
            "<li><b>Concrete pours &amp; curing</b> — protect fresh work from rain and from rapid "
            "drying; keep traffic off cemented layers at least 7 days (PMDM 1999 curing; "
            "SSRW 2000 Cl. 3803).</li>"
            "<li><b>Drainage/structures</b> — Zanzibar is in the PMDM WET zone, so moisture "
            "ingress is the primary risk; favour drainage works in wet spells "
            "(PMDM 1999 Table 2.1).</li>"
            "<li><b>Material haulage &amp; stockpiles</b> — cover stockpiles and expect haul-road "
            "softening; do not place a layer on one softened by excessive moisture "
            "(SSRW 2000 Cl. 3507).</li></ul>",
        "site_safety":
            "<ul><li><b>Crew safety</b> — stand down for lightning; manage heat stress at midday.</li>"
            "<li><b>Plant &amp; equipment</b> — reduced traction on wet subgrade; stage plant on "
            "firm ground.</li>"
            "<li><b>Excavation/slope stability</b> — batter and dewater open excavations ahead of "
            "rain.</li>"
            "<li><b>Flooding of works</b> — clear low-lying excavations and protect fresh layers.</li>"
            "<li><b>Haul-road condition</b> — inspect and maintain access after each wet period.</li></ul>",
    }


# ── HTML rendering (deterministic, exact spec) ────────────────────────────────
def esc(x):
    return html.escape(str(x if x is not None else ""), quote=False)


def _th(txt):
    return (f'<th style="background:{CLR["navy"]};color:{CLR["white"]};padding:8px 10px;'
            f'text-align:left;font-size:12px;border:1px solid {CLR["border"]}">{txt}</th>')


def _td(txt, i=0):
    bg = CLR["row_a"] if i % 2 == 0 else CLR["row_b"]
    return (f'<td style="background:{bg};padding:7px 10px;font-size:13px;'
            f'border:1px solid {CLR["border"]};color:{CLR["body"]}">{txt}</td>')


def section_title(t):
    return (f'<h2 style="color:{CLR["navy"]};border-bottom:2px solid {CLR["navy"]};'
            f'padding-bottom:6px;margin:26px 0 14px;font-size:18px">{t}</h2>')


def rfis_by_code(scope):
    """Group the in-scope RFIs by road code so each road card can list its works.

    RFIs carry a road CODE only, not an alignment, so when a code has several
    alignments (AL1/AL2/...) the same works are shown on each of that code's
    cards — they are the works recorded against that road.
    """
    by = {}
    for x in scope.get("items", []):
        c = (x.get("code") or "").strip().upper()
        if c:
            by.setdefault(c, []).append(x)
    for c in by:
        by[c].sort(key=lambda x: x["d"], reverse=True)
    return by


def rfi_block(code, by_code):
    """The 'RFI Works (last 3 days)' list shown inside a road card."""
    items = by_code.get((code or "").upper(), [])
    if not items:
        return (f'<div style="margin:4px 0 12px;font-size:12.5px;color:#6b7280">'
                f'<b>RFI Works (last 3 days):</b> none recorded against this road.</div>')
    lines = []
    for x in items:
        d = x["d"].strftime("%d-%m-%Y")
        act = (x.get("act") or x.get("desc") or "").strip().upper()
        desc = (x.get("desc") or "").strip()
        stn = (x.get("stn") or "").strip()
        parts = [f'{esc(d)} | RFI {esc(x.get("rfi") or "N/A")}: <b>{esc(act)}</b>']
        if desc and desc.upper() != act:
            parts.append(f' ({esc(desc)})')
        if stn:
            parts.append(f', Ch {esc(stn)}')
        lines.append(f'<li style="margin-bottom:3px">{"".join(parts)}</li>')
    return (f'<div style="margin:4px 0 12px;font-size:12.5px;line-height:1.55">'
            f'<div style="font-weight:700;color:{CLR["navy"]};margin-bottom:5px">'
            f'RFI Works (last 3 days):</div>'
            f'<ul style="margin:0;padding-left:18px">{"".join(lines)}</ul></div>')


def render_report(island, rows, scope, narr, provider, now):
    cfg = ISLANDS[island]
    L = cfg["label"]
    modelled = any(r["w"].get("modelled") for r in rows)
    src = ("MSN Weather (hourly values modelled from the daily forecast)"
           if provider == "msn" else "Open-Meteo hourly forecast")
    by_code = rfis_by_code(scope)
    codes_line = ", ".join(scope.get("codes", [])[:12]) + (
        " …" if len(scope.get("codes", [])) > 12 else "")
    P = [f'<!DOCTYPE html><html><head><meta charset="utf-8">'
         f'<meta name="viewport" content="width=device-width,initial-scale=1">'
         f'<title>IRIS-LUPTAN JV Road Weather Advisory — {esc(L)}</title></head>'
         f'<body style="margin:0;padding:0;background:{CLR["page"]};'
         f'font-family:{FONT};color:{CLR["body"]}">'
         f'<div style="max-width:1000px;margin:0 auto;padding:18px">']

    # SECTION 1 — header banner
    P.append(f'<div style="background:{CLR["navy"]};color:{CLR["white"]};padding:20px 24px;'
             f'border-radius:10px">'
             f'<div style="font-size:21px;font-weight:700;letter-spacing:.3px">'
             f'IRIS-LUPTAN JV ROADS_WEATHER ARTIFICIAL INTELLIGENCY ADVISORY</div>'
             f'<div style="font-size:15px;margin-top:8px">Island: <b>{esc(L)}</b></div>'
             f'<div style="font-size:13px;opacity:.92;margin-top:6px">'
             f'Data source: {esc(src)} &nbsp;·&nbsp; Report generated: '
             f'{now.strftime("%d-%m-%Y %H:%M")} EAT<br>'
             f'Coverage: {len(rows)} {esc(L.split()[0].title())} RFI road alignment(s)'
             + (f' ({esc(codes_line)})' if codes_line.strip() else '')
             + f' &nbsp;·&nbsp; {scope["count"]} site-work RFI(s) in the last '
             f'{scope.get("days", 3)} days<br>'
             f'Issued by IRIS-LUPTAN JV Weather Intelligence</div></div>')

    # SECTION 2 — executive summary + 3-7 day outlook
    P.append(section_title("SECTION 2 — EXECUTIVE SUMMARY"))
    P.append(f'<div style="background:{CLR["card"]};padding:16px 20px;border-radius:10px;'
             f'border:1px solid {CLR["border"]};font-size:14px;line-height:1.65">'
             f'{narr["executive_summary"]}</div>')
    if rows:
        P.append('<div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse;'
                 'margin-top:14px"><tr>'
                 + "".join(_th(h) for h in ["Day", "Date", "Condition", "High °C", "Low °C",
                                            "Precip %", "Wind km/h", "Status"]) + "</tr>")
        for i, d in enumerate(rows[0]["w"]["daily"]):
            P.append("<tr>" + _td(esc(d["day"]), i) + _td(esc(d["date"]), i)
                     + _td(f'{d["icon"]} {esc(d["text"])}', i) + _td(esc(d["high"]), i)
                     + _td(esc(d["low"]), i) + _td(f'💧 {esc(d["precip"])}%', i)
                     + _td(f'💨 {esc(d["wind"])} km/h', i) + _td(badge(d["status"]), i) + "</tr>")
        P.append("</table></div>")

    # SECTION 3 — per-road detail
    P.append(section_title("SECTION 3 — PER-ROAD DETAIL"))
    if not rows:
        P.append(f'<div style="background:{CLR["card"]};padding:18px;border-radius:10px;'
                 f'border:1px solid {CLR["border"]}">No roads in scope — there were no '
                 f'site-work RFIs in the last {scope.get("days", 3)} days.</div>')
    for r in rows:
        hs = r["w"]["hourly"]
        dom = worst([h["status"] for h in hs])
        dom_icon = next((h["icon"] for h in hs if h["status"] == dom), hs[0]["icon"])
        cu = r["w"]["current"]
        P.append(f'<div style="background:{CLR["card"]};border:1px solid {CLR["border"]};'
                 f'border-radius:10px;padding:16px 18px;margin-bottom:14px">'
                 f'<div style="font-weight:700;color:{CLR["navy"]};font-size:15px">'
                 f'{dom_icon} {esc(r["road"]["name"])}</div>'
                 f'<div style="margin:9px 0 12px;font-size:13px">Current: '
                 f'{cu["icon"]} {esc(cu["text"])} &nbsp;·&nbsp; {esc(cu["temp"])} °C '
                 f'&nbsp;·&nbsp; 💧 {esc(cu["humidity"])}% &nbsp;·&nbsp; '
                 f'💨 {esc(cu["wind"])} km/h &nbsp; {badge(cu["status"])}</div>'
                 + rfi_block(r["road"]["name"].split()[0], by_code)
                 + f'<div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse"><tr>'
                 + _th("Time") + _th("Temp °C") + _th(f"Condition {dom_icon}")
                 + _th("Humidity %") + _th("Wind km/h") + _th("Precip %") + _th("Status")
                 + "</tr>")
        for i, h in enumerate(hs):
            P.append("<tr>" + _td(f'{h["hour"]:02d}:00', i) + _td(esc(h["temp"]), i)
                     + _td(f'{h["icon"]} {esc(h["text"])}', i) + _td(f'💧 {esc(h["humidity"])}%', i)
                     + _td(f'💨 {esc(h["wind"])} km/h', i) + _td(f'💧 {esc(h["precip"])}%', i)
                     + _td(badge(h["status"]), i) + "</tr>")
        P.append("</table></div></div>")

    # SECTION 4 — risk matrix & construction advisories
    P.append(section_title("SECTION 4 — RISK MATRIX &amp; CONSTRUCTION ADVISORIES"))
    P.append(f'<div style="background:{CLR["card"]};border:1px solid {CLR["border"]};'
             f'border-radius:10px;padding:16px 18px">')
    for s in STATUS_ORDER:
        grp = [r for r in rows if r["status"] == s]
        if grp:
            P.append(f'<div style="margin-bottom:10px;font-size:13px">{badge(s)} &nbsp;'
                     + ", ".join(esc(r["road"]["name"]) for r in grp) + "</div>")
    P.append(f'<h3 style="color:{CLR["navy"]};font-size:14px;margin:18px 0 6px">'
             f'(a) WORKS AFFECTED TODAY</h3>'
             f'<div style="font-size:13.5px;line-height:1.7">{narr["works_affected"]}</div>'
             f'<h3 style="color:{CLR["navy"]};font-size:14px;margin:18px 0 6px">'
             f'(b) SITE SAFETY &amp; ACCESS</h3>'
             f'<div style="font-size:13.5px;line-height:1.7">{narr["site_safety"]}</div></div>')

    # SECTION 5 — footer + status legend
    P.append(section_title("SECTION 5 — FOOTER"))
    P.append(f'<div style="background:{CLR["card"]};border:1px solid {CLR["border"]};'
             f'border-radius:10px;padding:16px 18px;font-size:12.5px;line-height:1.7">'
             f'<p style="margin:0 0 10px">All construction advice follows Tanzania\'s '
             f'<b>Pavement and Materials Design Manual 1999 (PMDM 1999)</b>, adopted by Zanzibar, '
             f'with clause citations from <b>SSRW 2000</b>. Weather data is a forecast and must be '
             f'confirmed on site before committing plant or materials.'
             + (' Hourly values from this provider are modelled from the daily forecast, not '
                'measured hourly observations.' if modelled else '')
             + f'</p><p style="margin:0 0 14px">Next report: tomorrow {cfg["next"]} EAT '
             f'&nbsp;·&nbsp; <b>IRIS-LUPTAN JV</b> Weather Intelligence</p>'
             f'<div style="font-weight:700;color:{CLR["navy"]};margin-bottom:8px">'
             f'STATUS LEGEND</div>')
    for s in STATUS_ORDER[::-1]:
        P.append(f'<div style="margin-bottom:6px">{badge(s)} &nbsp;{STATUS_DESC[s]}</div>')
    P.append("</div></div></body></html>")
    return "".join(P)


# ── Email ─────────────────────────────────────────────────────────────────────
def load_recipients(island):
    """Recipients come from WEATHER_RECIPIENTS_JSON when set (CI secret), else
    from weather_assets/recipients.json. Addresses are never stored in code, and
    in CI they stay in the secret rather than the repository."""
    raw = os.environ.get("WEATHER_RECIPIENTS_JSON", "").strip()
    if raw:
        try:
            cfg = json.loads(raw)
        except json.JSONDecodeError as e:
            raise RuntimeError(f"WEATHER_RECIPIENTS_JSON is not valid JSON: {e}")
        src = "WEATHER_RECIPIENTS_JSON"
    else:
        if not os.path.exists(RECIPIENTS_FILE):
            raise RuntimeError(
                f"no recipients: set WEATHER_RECIPIENTS_JSON or create {RECIPIENTS_FILE}")
        cfg = json.load(open(RECIPIENTS_FILE, encoding="utf-8"))
        src = RECIPIENTS_FILE
    to = (cfg.get(island) or {}).get("to") or []
    if not to:
        raise RuntimeError(f"no recipients configured for '{island}' in {src}")
    return to


def with_dashboard_link(doc):
    """Add a button to the private dashboard (DASHBOARD_URL) to the emailed copy."""
    url = os.environ.get("DASHBOARD_URL", "").strip()
    tail = "</div></div></body></html>"
    if not url or not doc.endswith(tail):
        return doc
    link = ('<div style="text-align:center;margin:20px 0 6px">'
            f'<a href="{html.escape(url, quote=True)}" style="display:inline-block;'
            f'background:{CLR["navy"]};color:#ffffff;text-decoration:none;font-weight:700;'
            'font-size:14px;padding:12px 26px;border-radius:8px">'
            'Open the live weather dashboard</a>'
            '<div style="font-size:12px;color:#6b7280;margin-top:8px">'
            'Sign in with the email address that received this advisory; '
            'a one-time code will be sent to you.</div></div>')
    return doc[:-len(tail)] + "</div>" + link + "</div></body></html>"


def send_email(subject, to_list, body_html):
    user = os.environ.get("GMAIL_USER", "").strip()
    pw = os.environ.get("GMAIL_APP_PASSWORD", "").replace(" ", "")
    if not user or not pw:
        raise RuntimeError("GMAIL_USER and GMAIL_APP_PASSWORD must be set to send mail")
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = user
    msg["To"] = ", ".join(to_list)
    msg.attach(MIMEText("This advisory requires an HTML-capable mail client.", "plain", "utf-8"))
    msg.attach(MIMEText(body_html, "html", "utf-8"))
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=180) as s:
        s.login(user, pw)
        s.sendmail(user, to_list, msg.as_string())


# ── Per-island run ────────────────────────────────────────────────────────────
def run_island(island, args, svc, clauses):
    cfg = ISLANDS[island]
    now = datetime.now(EAT)
    log(f"\n{'=' * 70}\n  {cfg['label']}  —  {now:%Y-%m-%d %H:%M} EAT\n{'=' * 70}")

    log("[1] Reading RFI_DATA tabs …")
    rfi_rows = []
    for ssid, tab in cfg["sheets"]:
        got = read_tab(svc, ssid, tab)
        log(f"    {tab}: {len(got)} rows")
        rfi_rows += got

    log(f"[2] Scoping to site works of the last {args.days} day(s) …")
    scope = scope_rfis(rfi_rows, days=args.days, now=now)
    scope["days"] = args.days
    log(f"    {scope['count']} RFI(s) · {len(scope['codes'])} road code(s): "
        f"{', '.join(scope['codes'][:14])}{' …' if len(scope['codes']) > 14 else ''}")

    log("[3] Resolving coordinates …")
    roads = roads_in_scope(island, scope["codes"])
    if args.limit:
        roads = roads[: args.limit]
    surveyed = sum(1 for r in roads if r.get("coord_src") == "survey")
    log(f"    {len(roads)} alignment(s) in scope  ({surveyed} survey, "
        f"{len(roads) - surveyed} master coords)")
    missing = [c for c in scope["codes"]
               if not any(r["name"].split()[0].upper() == c.upper() for r in roads)]
    if missing:
        log(f"    ! no coordinates for: {', '.join(missing)}")

    log(f"[4] Fetching weather ({args.provider}) …")
    results = []
    if roads:
        with ThreadPoolExecutor(max_workers=6) as ex:
            for road, w, err in ex.map(lambda r: fetch_weather(r, args.provider), roads):
                if w:
                    results.append({"road": road, "w": w,
                                    "status": worst([h["status"] for h in w["hourly"]])})
                else:
                    log(f"    ! {road['name'][:48]}: {err}")
    results.sort(key=lambda r: (STATUS_ORDER.index(r["status"]), r["road"]["name"]))
    log(f"    {len(results)}/{len(roads)} road(s) retrieved")

    log("[5] Narrative …")
    narr = None if args.no_ai else ai_narrative(island, results, scope, clauses)
    if narr:
        log("    Claude narrative generated")
    else:
        narr = fallback_narrative(results)
        log("    using built-in narrative")

    log("[6] Rendering …")
    doc = render_report(island, results, scope, narr, args.provider, now)
    out = args.out or os.path.join(BASE_DIR, "logs", f"weather_{island}_{now:%Y%m%d}.html")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with open(out, "w", encoding="utf-8") as f:
        f.write(doc)
    log(f"    {len(doc):,} bytes -> {out}")

    if args.publish:
        from weather_publish import publish_advisory
        rel = publish_advisory(args.publish, island, cfg["label"], now, doc, results, scope,
                               worst([r["status"] for r in results]) if results else None)
        log(f"    published -> {rel}")

    subject = cfg["subject"].format(date=now.strftime("%d-%m-%Y"))
    if args.dry_run:
        log(f"[7] DRY RUN — not sending.  Subject: {subject}")
    else:
        to = load_recipients(island)
        log(f"[7] Sending to {len(to)} recipient(s) …")
        send_email(subject, to, with_dashboard_link(doc))
        log("    sent")
    return len(results)


def main():
    ap = argparse.ArgumentParser(description="IRIS Road Weather Advisory")
    ap.add_argument("--island", default="both", choices=["unguja", "pemba", "both"])
    ap.add_argument("--dry-run", action="store_true", help="render + save, do not email")
    ap.add_argument("--no-ai", action="store_true", help="skip Claude, use built-in narrative")
    ap.add_argument("--provider", default=os.environ.get("WEATHER_PROVIDER", "openmeteo"),
                    choices=["openmeteo", "msn"])
    ap.add_argument("--days", type=int, default=3, help="RFI look-back window")
    ap.add_argument("--limit", type=int, default=0, help="cap roads (testing)")
    ap.add_argument("--out", default=None, help="output HTML path (single island)")
    ap.add_argument("--publish", default=None, metavar="DIR",
                    help="also write the advisory and index.json into this site folder")
    args = ap.parse_args()

    log("=" * 70)
    log("  IRIS ROAD WEATHER ADVISORY")
    log("=" * 70)
    clauses = load_clauses()
    log(f"  Spec clauses loaded: {len(clauses)}"
        + (f"  ({', '.join(c['manual'] + ' ' + c['clause'] for c in clauses[:5])} …)"
           if clauses else ""))
    svc = sheets_service()
    log("  Google Sheets: authenticated")

    islands = ["unguja", "pemba"] if args.island == "both" else [args.island]
    total = 0
    failed = 0
    for isl in islands:
        try:
            total += run_island(isl, args, svc, clauses)
        except Exception as e:
            failed += 1
            log(f"\n  !! {isl} FAILED: {type(e).__name__}: {e}")
    log(f"\n  DONE — {total} road report(s) across {len(islands)} island(s)"
        + (f", {failed} island(s) failed" if failed else "") + "\n")
    return 1 if failed else 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        log("\nInterrupted.")
        sys.exit(130)

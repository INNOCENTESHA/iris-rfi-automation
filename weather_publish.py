"""Dashboard publishing for the IRIS Road Weather Advisory.

Writes the day's advisory and a summary index into a site folder:
    advisories/<island>/<YYYY-MM-DD>.html
    advisories/index.json
The folder is then uploaded to private storage by the workflow.
"""

import json
import os


def publish_advisory(dest, island, label, now, doc, results, scope, worst_status):
    """Write the advisory and update advisories/index.json under `dest`."""
    date = now.strftime("%Y-%m-%d")
    rel = f"advisories/{island}/{date}.html"
    path = os.path.join(dest, *rel.split("/"))
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write(doc)
    tally = {}
    for r in results:
        tally[r["status"]] = tally.get(r["status"], 0) + 1
    entry = {"island": island, "label": label, "date": date,
             "generated": now.strftime("%Y-%m-%d %H:%M EAT"), "file": rel,
             "roads": len(results), "rfis": scope["count"], "codes": scope["codes"],
             "tally": tally,
             "worst": worst_status,
             "road_list": [], "outlook": []}
    for r in results:
        hs = (r.get("w") or {}).get("hourly") or []
        temps = [h["temp"] for h in hs if h.get("temp") is not None]
        dom = next((h for h in hs if h["status"] == r["status"]), hs[0] if hs else {})
        entry["road_list"].append({
            "name": r["road"]["name"], "status": r["status"], "icon": dom.get("icon", ""),
            "text": dom.get("text", ""),
            "precip": max([h.get("precip") or 0 for h in hs], default=0),
            "wind": max([h.get("wind") or 0 for h in hs], default=0),
            "tmin": min(temps) if temps else None, "tmax": max(temps) if temps else None})
    if results:
        for d in ((results[0].get("w") or {}).get("daily") or [])[:7]:
            entry["outlook"].append({k: d.get(k) for k in
                                     ("day", "date", "icon", "text", "high", "low",
                                      "precip", "wind", "status")})
    idx_path = os.path.join(dest, "advisories", "index.json")
    try:
        idx = json.load(open(idx_path, encoding="utf-8"))
    except Exception:
        idx = []
    idx = [e for e in idx if not (e.get("island") == island and e.get("date") == date)]
    idx.append(entry)
    idx.sort(key=lambda e: (e["date"], e["island"]), reverse=True)
    with open(idx_path, "w", encoding="utf-8") as f:
        json.dump(idx[:800], f, indent=1, ensure_ascii=False)
    return rel

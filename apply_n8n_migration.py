"""
IRIS n8n Workflow Migration Script
===================================
Run this once you have a fresh n8n API key from:
  https://innocentesha.app.n8n.cloud/settings/api

Usage:
  python apply_n8n_migration.py <YOUR_N8N_API_KEY>

What it does:
  1. Fetches current workflow from n8n cloud API
  2. Replaces the 3 dead VPS SSH nodes with GitHub Actions / Wait / No-op nodes
  3. Updates the 2 HTTP Retry nodes to point at GitHub Actions dispatch API
  4. PUTs the modified workflow back
  5. Activates the workflow
"""

import sys
import json
import copy
import urllib.request
import urllib.error

# ── Config ─────────────────────────────────────────────────────────────────────
N8N_BASE    = "https://innocentesha.app.n8n.cloud/api/v1"
WORKFLOW_ID = "OOaOngWoipzNxoA9"
GH_URL      = "https://api.github.com/repos/INNOCENTESHA/iris-rfi-automation/dispatches"
GH_PAT      = "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"

if len(sys.argv) < 2:
    print("Usage: python apply_n8n_migration.py <N8N_API_KEY>")
    sys.exit(1)

N8N_KEY = sys.argv[1].strip()

# ── HTTP helper ────────────────────────────────────────────────────────────────
def n8n_request(method, path, body=None):
    url = f"{N8N_BASE}{path}"
    data = json.dumps(body).encode() if body is not None else None
    req = urllib.request.Request(url, data=data, method=method)
    req.add_header("X-N8N-API-KEY", N8N_KEY)
    req.add_header("Content-Type", "application/json")
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return resp.status, json.loads(resp.read())
    except urllib.error.HTTPError as e:
        return e.code, json.loads(e.read())

# ── Step 1: Fetch workflow ─────────────────────────────────────────────────────
print(f"Fetching workflow {WORKFLOW_ID}...")
status, wf = n8n_request("GET", f"/workflows/{WORKFLOW_ID}")
if status != 200:
    print(f"  FAILED {status}: {wf}")
    sys.exit(1)
print(f"  OK — {len(wf['nodes'])} nodes found")

# ── Step 2: Modify nodes ───────────────────────────────────────────────────────
GH_PARAMS = {
    "curlImport": "",
    "method": "POST",
    "url": GH_URL,
    "authentication": "genericCredentialType",
    "genericAuthType": "httpHeaderAuth",
    "provideSslCertificates": False,
    "sendQuery": False,
    "sendHeaders": True,
    "headerParameters": {
        "parameters": [
            {"name": "Authorization", "value": f"Bearer {GH_PAT}"},
            {"name": "Accept",        "value": "application/vnd.github+json"},
            {"name": "X-GitHub-Api-Version", "value": "2022-11-28"}
        ]
    },
    "sendBody": True,
    "contentType": "json",
    "specifyBody": "json",
    "jsonBody": '{"event_type":"run-iris","client_payload":{"mode":"full-scan"}}',
    "options": {}
}

name_map = {}
for node in wf["nodes"]:
    if node["name"] == "SSH: Trigger IRIS Run":
        print("  Replacing SSH: Trigger IRIS Run → HTTP: Trigger GitHub Actions")
        node["name"] = "HTTP: Trigger GitHub Actions"
        node["type"] = "n8n-nodes-base.httpRequest"
        node["typeVersion"] = 4.2
        node["parameters"] = copy.deepcopy(GH_PARAMS)
        node.pop("credentials", None)
        name_map["SSH: Trigger IRIS Run"] = "HTTP: Trigger GitHub Actions"

    elif node["name"] == "SSH: Restart VPS Service":
        print("  Replacing SSH: Restart VPS Service → Wait: 60s For GHA Start")
        node["name"] = "Wait: 60s For GHA Start"
        node["type"] = "n8n-nodes-base.wait"
        node["typeVersion"] = 1.1
        node["parameters"] = {"resume": "timeInterval", "amount": 60, "unit": "seconds"}
        node.pop("credentials", None)
        name_map["SSH: Restart VPS Service"] = "Wait: 60s For GHA Start"

    elif node["name"] == "SSH: Read VPS Logs":
        print("  Replacing SSH: Read VPS Logs → No Op: Skip Log Read")
        node["name"] = "No Op: Skip Log Read"
        node["type"] = "n8n-nodes-base.noOp"
        node["typeVersion"] = 1
        node["parameters"] = {}
        node.pop("credentials", None)
        name_map["SSH: Read VPS Logs"] = "No Op: Skip Log Read"

    elif node["name"] == "HTTP Retry: Post-Restart":
        print("  Updating HTTP Retry: Post-Restart → GitHub Actions URL")
        node["parameters"].update(GH_PARAMS)

    elif node["name"] == "HTTP Retry: Post-Server-Error":
        print("  Updating HTTP Retry: Post-Server-Error → GitHub Actions URL")
        node["parameters"].update(GH_PARAMS)

# Fix connections (rename keys for renamed source nodes)
new_conns = {}
for k, v in wf.get("connections", {}).items():
    new_conns[name_map.get(k, k)] = v
wf["connections"] = new_conns

# ── Step 3: PUT (only send allowed fields) ─────────────────────────────────────
# Strip read-only fields; keep only what the API PUT endpoint accepts
SAFE_SETTINGS_KEYS = {"executionOrder", "callerPolicy", "errorWorkflow",
                      "saveManualExecutions", "saveExecutionProgress",
                      "saveDataSuccessExecution", "saveDataErrorExecution", "timezone"}

safe_settings = {k: v for k, v in wf.get("settings", {}).items() if k in SAFE_SETTINGS_KEYS}

payload = {
    "name":        wf["name"],
    "nodes":       wf["nodes"],
    "connections": wf["connections"],
    "settings":    safe_settings,
    "staticData":  wf.get("staticData") or None,
}

print(f"\nPUTting modified workflow...")
status, result = n8n_request("PUT", f"/workflows/{WORKFLOW_ID}", payload)
if status == 200:
    print(f"  SUCCESS — workflow saved (versionId: {result.get('versionId', '?')})")
else:
    print(f"  FAILED {status}: {json.dumps(result)[:400]}")
    sys.exit(1)

# ── Step 4: Activate ──────────────────────────────────────────────────────────
print("\nActivating workflow...")
status, result = n8n_request("POST", f"/workflows/{WORKFLOW_ID}/activate")
if status == 200:
    print("  Workflow activated!")
else:
    print(f"  Activation status {status}: {result}")

print("\nDone. Summary of changes:")
for old, new in name_map.items():
    print(f"  {old} → {new}")
print("  HTTP Retry: Post-Restart → GitHub Actions URL")
print("  HTTP Retry: Post-Server-Error → GitHub Actions URL")

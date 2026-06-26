"""
IRIS Trigger Server  (v1 — May 2026)
=====================================
A lightweight HTTP server that listens for n8n's Schedule Trigger and
starts iris_launcher.py — allowing n8n Cloud to fully orchestrate the
IRIS pipeline even though the Python scripts run locally on this machine.

HOW IT WORKS:
  1. This server runs permanently on this PC (port 8765)
  2. n8n Schedule fires at 17:00 → n8n sends POST to /run-iris
  3. This server spawns iris_launcher.py in a new terminal window
  4. iris_launcher.py runs all 3 scripts and POSTs results back to n8n
  5. n8n logs to Google Sheets and sends the Gmail summary

HOW TO START:
  python C:\\AI-Projects\\iris_trigger_server.py

  Or double-click:  run_iris_server.bat

TO EXPOSE TO N8N CLOUD (required — run once in a separate window):
  ngrok http 8765
  ← copy the https://xxxx.ngrok-free.app URL into n8n HTTP Request node

MANUAL TRIGGER (test without n8n):
  curl -X POST http://localhost:8765/run-iris
  or open:  http://localhost:8765/status
"""

import sys, os, json, subprocess, threading, datetime
from http.server import HTTPServer, BaseHTTPRequestHandler

SCRIPT_DIR   = r"C:\AI-Projects"
LAUNCHER     = os.path.join(SCRIPT_DIR, "iris_launcher.py")
PORT         = 8765
LOG_FILE     = os.path.join(SCRIPT_DIR, "iris_trigger_log.txt")

_lock        = threading.Lock()
_last_run    = None          # datetime of last trigger
_running     = False         # True while iris_launcher is active
_run_count   = 0


def _log(msg: str):
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line, flush=True)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def _spawn_launcher():
    """Run iris_launcher.py in a new visible terminal window."""
    global _running, _last_run, _run_count
    with _lock:
        if _running:
            _log("SKIP — launcher already running")
            return
        _running  = True
        _last_run = datetime.datetime.now()
        _run_count += 1

    _log(f"TRIGGER #{_run_count} — starting iris_launcher.py")
    try:
        # Open a new CMD window so the Rich dashboard is visible
        subprocess.Popen(
            ["cmd", "/c", "start", "cmd", "/k",
             sys.executable, LAUNCHER],
            cwd=SCRIPT_DIR,
            creationflags=subprocess.CREATE_NEW_CONSOLE,
        )
    except Exception as e:
        _log(f"ERROR spawning launcher: {e}")
    finally:
        # Mark as not-running after a short delay so rapid double-triggers
        # are still blocked, but a real second run can happen.
        def _clear():
            import time; time.sleep(30)
            global _running
            _running = False
        threading.Thread(target=_clear, daemon=True).start()


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        _log(f"HTTP {fmt % args}")

    def _send(self, code: int, body: dict):
        data = json.dumps(body).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", len(data))
        self.end_headers()
        self.wfile.write(data)

    # ── POST /run-iris  ← n8n calls this ──────────────────────────────────────
    def do_POST(self):
        if self.path.rstrip("/") == "/run-iris":
            threading.Thread(target=_spawn_launcher, daemon=True).start()
            self._send(200, {"status": "accepted", "message": "iris_launcher.py started"})
        else:
            self._send(404, {"error": "unknown path"})

    # ── GET /status  ← health check ───────────────────────────────────────────
    def do_GET(self):
        if self.path.rstrip("/") == "/status":
            self._send(200, {
                "status":    "running" if _running else "idle",
                "last_run":  _last_run.isoformat() if _last_run else None,
                "run_count": _run_count,
                "launcher":  LAUNCHER,
            })
        else:
            self._send(404, {"error": "unknown path"})


if __name__ == "__main__":
    _log(f"IRIS Trigger Server starting on port {PORT}")
    _log(f"Launcher path : {LAUNCHER}")
    _log(f"Log file      : {LOG_FILE}")
    _log("Waiting for n8n trigger at POST /run-iris ...")
    _log("Health check  : GET  http://localhost:{PORT}/status")
    print()

    server = HTTPServer(("0.0.0.0", PORT), Handler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        _log("Server stopped by user.")

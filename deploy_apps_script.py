#!/usr/bin/env python3
"""
deploy_apps_script.py — push local Apps Script code and update the EXISTING
web app deployment in place (the /exec URL in .env never changes).

Flow:
  1. reads .env  ->  appscript_url  (…/macros/s/<DEPLOYMENT_ID>/exec)
  2. clasp push                  (uploads Code.js + appsscript.json)
  3. clasp create-version "<msg>"  (immutable snapshot)
  4. clasp update-deployment <DEPLOYMENT_ID from .env> --versionNumber <N>
       -> the same deployment now serves the new version; URL unchanged
  5. GET <appscript_url>          (health check; retries while Google propagates)

Usage:
  python deploy_apps_script.py                # full deploy
  python deploy_apps_script.py --dry-run      # push only what WOULD happen
  python deploy_apps_script.py -m "message"   # version description

Rollback (if a deploy misbehaves):
  clasp update-deployment <DEPLOYMENT_ID> --versionNumber 3
  (each version stays listed under: clasp list-versions)

Requires: clasp (npm i -g @google/clasp) and an active `clasp login`.
"""
from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
import time
import urllib.request
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
ENV_FILE = SCRIPT_DIR / ".env"
DEFAULT_DEPLOYMENT_DESC = "update from local Code.js"


def log(msg: str) -> None:
    print(msg, file=sys.stderr, flush=True)


def setting(key: str) -> str:
    """Read 'key :- value' from .env (same format run_billing.py uses)."""
    if not ENV_FILE.exists():
        return ""
    for line in ENV_FILE.read_text(encoding="utf-8", errors="ignore").splitlines():
        m = re.match(r"\s*" + re.escape(key) + r"\s*:-\s*(.+)", line)
        if m:
            return m.group(1).strip()
    return ""


def deployment_id_from_url(url: str) -> str:
    """https://script.google.com/macros/s/AKfycb…/exec -> AKfycb…"""
    m = re.search(r"/macros/s/([^/]+)/", url or "")
    if not m:
        raise SystemExit("ERROR: .env appscript_url is not a valid web app /exec URL")
    return m.group(1)


def run_clasp(args: list[str], *, capture: bool = True) -> subprocess.CompletedProcess:
    exe = shutil.which("clasp")               # npm installs clasp.cmd on Windows
    if not exe:
        raise SystemExit("ERROR: clasp not found — npm install -g @google/clasp")
    cmd = [exe, *args]
    log(f"$ {' '.join(cmd)}")
    return subprocess.run(
        cmd, cwd=str(SCRIPT_DIR), capture_output=capture, text=True,
        encoding="utf-8", errors="replace",
    )


def clasp_out(args: list[str]) -> str:
    cp = run_clasp(args)
    if cp.returncode != 0:
        raise SystemExit(f"ERROR: clasp {' '.join(args)} failed:\n{cp.stdout}\n{cp.stderr}")
    return (cp.stdout or "") + (cp.stderr or "")


def health_check(url: str, tries: int = 6, wait_s: int = 10) -> None:
    """GET /exec until doGet answers; Google can lag a little after redeploy."""
    for attempt in range(1, tries + 1):
        try:
            with urllib.request.urlopen(url, timeout=60) as resp:
                body = json.loads(resp.read().decode("utf-8"))
            if body.get("ok"):
                log(f"[health] OK: {body}")
                return
            log(f"[health] unexpected body: {body}")
        except Exception as exc:
            log(f"[health] attempt {attempt}/{tries} failed: {exc}")
        if attempt < tries:
            time.sleep(wait_s)
    raise SystemExit("ERROR: web app did not answer the health check after redeploy")


def main() -> int:
    ap = argparse.ArgumentParser(description="Push Code.js and update the existing web app deployment (URL preserved).")
    ap.add_argument("-m", "--message", default=DEFAULT_DEPLOYMENT_DESC, help="version description")
    ap.add_argument("--dry-run", action="store_true", help="show the plan and push nothing")
    ap.add_argument("--skip-push", action="store_true", help="skip clasp push (code already on remote)")
    args = ap.parse_args()

    url = setting("appscript_url") or ""
    if not url:
        raise SystemExit("ERROR: .env has no 'appscript_url :-' line")
    dep_id = deployment_id_from_url(url)
    log(f"[plan] deployment to update (from .env): {dep_id}")
    log(f"[plan] webhook URL stays      : {url}")

    # 0. sanity: no local conflict, files exist
    if not (SCRIPT_DIR / "Code.js").exists():
        raise SystemExit("ERROR: Code.js not found next to this script")
    st = run_clasp(["status"])
    if st.returncode != 0:
        raise SystemExit(f"ERROR: clasp status failed ({st.stdout} {st.stderr}) — resolve conflicts first")
    log("[plan] files that will be pushed:\n" + (st.stdout or "").rstrip())

    if args.dry_run:
        log("[dry-run] nothing was pushed or deployed")
        return 0

    # 1. push
    if not args.skip_push:
        clasp_out(["push"])

    # 2. immutable version
    ver_out = clasp_out(["create-version", args.message, "--json"])
    try:
        ver = json.loads(ver_out.strip().splitlines()[-1])  # tolerate clasp chatter
        version_number = int(ver.get("versionNumber") or ver.get("version") or 0)
    except (ValueError, json.JSONDecodeError):
        # fallback: last "N - description" line of list-versions
        lv = clasp_out(["list-versions"])
        nums = [int(x) for x in re.findall(r"^(\d+)\s*-", lv, re.M)]
        if not nums:
            raise SystemExit(f"ERROR: could not determine new version number\n{ver_out}")
        version_number = max(nums)
    log(f"[version] created version {version_number}")

    # 3. point the EXISTING deployment at it (URL unchanged)
    clasp_out(["update-deployment", dep_id, "--versionNumber", str(version_number),
               "-d", args.message, "--json"])
    log(f"[deploy] deployment {dep_id} now serves version {version_number}")

    # 4. prove the webhook is alive
    health_check(url)

    print(json.dumps({"ok": True, "deployment_id": dep_id,
                      "version": version_number, "url": url}, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())

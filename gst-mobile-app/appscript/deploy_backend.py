#!/usr/bin/env python
"""Deploy the GST backend (Code.js) straight from the PC — no copy-paste.

Replaces the manual flow (paste Code.js in the Apps Script editor -> Deploy ->
Manage deployments -> New version) with one command:

    python deploy_backend.py login     # ONE-TIME: approve in the browser
    python deploy_backend.py push      # push Code.js + new version, SAME
                                       # deployment id -> live immediately

What "push" does (identical to the web UI flow):
  1. uploads Code.js into the Apps Script project (content update)
  2. creates a new VERSION
  3. points the EXISTING deployment (same id -> same /exec URL) at that
     version

Config lives in deploy.json next to this script:
  { "scriptId": "<Apps Script project id>", "deploymentId": "AKfycb..." }
The deploymentId is the path segment of the /exec URL. The Script ID is in
the Apps Script editor: Project Settings -> Script ID.

Auth uses clasp's public OAuth client (same one clasp itself uses) and a
one-time browser approval; the refresh token is saved to .deploy-token.json
(gitignored) so logins persist.
"""
import json
import os
import socket
import sys
import threading
import urllib.parse
import urllib.request
from http.server import BaseHTTPRequestHandler, HTTPServer

HERE = os.path.dirname(os.path.abspath(__file__))
CONFIG = os.path.join(HERE, 'deploy.json')
TOKENS = os.path.join(HERE, '.deploy-token.json')

CLIENT_ID = '1072944905499-vm2v2i5dvn0a0d2o4ca36i1vge8cvbn0.apps.googleusercontent.com'  # clasp public client
CLIENT_SECRET = 'v6V3fKV_zWU7iw1DrpO1rknX'
SCOPES = ['https://www.googleapis.com/auth/script.projects',
          'https://www.googleapis.com/auth/script.deployments']
API = 'https://script.googleapis.com/v1/projects'


def load_config():
    if not os.path.exists(CONFIG):
        sys.exit('deploy.json not found — create it with {"scriptId": "...", "deploymentId": "AKfycb..."}')
    cfg = json.load(open(CONFIG, encoding='utf-8'))
    if not cfg.get('scriptId') or 'PASTE' in cfg['scriptId']:
        sys.exit('put the Script ID into deploy.json first (Apps Script editor -> Project Settings -> Script ID)')
    return cfg


def load_token():
    if not os.path.exists(TOKENS):
        sys.exit('not logged in — run:  python deploy_backend.py login')
    return json.load(open(TOKENS, encoding='utf-8'))


def save_token(tok):
    json.dump(tok, open(TOKENS, 'w', encoding='utf-8'), indent=1)


def post_form(url, data):
    req = urllib.request.Request(url, data=urllib.parse.urlencode(data).encode(),
                                 headers={'Content-Type': 'application/x-www-form-urlencoded'})
    try:
        return json.load(urllib.request.urlopen(req, timeout=60))
    except urllib.error.HTTPError as e:
        sys.exit('token exchange failed: ' + e.read().decode('utf-8', 'replace')[:400])


def api(method, url, token, body=None):
    data = json.dumps(body).encode() if body is not None else None
    req = urllib.request.Request(url, data=data, method=method,
                                 headers={'Authorization': 'Bearer ' + token,
                                          'Content-Type': 'application/json'})
    try:
        resp = urllib.request.urlopen(req, timeout=120)
        raw = resp.read().decode('utf-8')
        return json.loads(raw) if raw else {}
    except urllib.error.HTTPError as e:
        raise SystemExit('%s failed: %s' % (method, e.read().decode('utf-8', 'replace')[:600]))


def access_token(tok):
    if tok.get('access_token') and tok.get('expiry', 0) > __import__('time').time() + 60:
        return tok['access_token']
    fresh = post_form('https://oauth2.googleapis.com/token', {
        'client_id': CLIENT_ID, 'client_secret': CLIENT_SECRET,
        'refresh_token': tok['refresh_token'], 'grant_type': 'refresh_token'})
    tok['access_token'] = fresh['access_token']
    tok['expiry'] = __import__('time').time() + fresh.get('expires_in', 3600) - 120
    save_token(tok)
    return tok['access_token']


# --------------------------------------------------------------- login ----
def cmd_login():
    """One-time: open the browser, approve, the local server catches the
    redirect and stores the refresh token. Approve with the Google account
    that OWNS the script (shreeganeshtelecom14)."""
    result = {}

    class Handler(BaseHTTPRequestHandler):
        def do_GET(self):
            q = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
            result['code'] = (q.get('code') or [None])[0]
            result['error'] = (q.get('error') or [None])[0]
            print('   <- callback received: %s' % ('code OK' if result['code'] else 'error=' + str(result['error'])), flush=True)
            self.send_response(200)
            self.send_header('Content-Type', 'text/html')
            self.end_headers()
            ok = result.get('code') and not result.get('error')
            self.wfile.write(('<h2>%s</h2><script>setTimeout(()=>window.close(),1500)</script>'
                              % ('✅ Logged in — you can close this tab.' if ok
                                 else '❌ Login failed: ' + str(result.get('error')))).encode())

        def log_message(self, *a):
            print('   <- request:', a[0] if a else '', flush=True)

    port = 5577
    server = HTTPServer(('localhost', port), Handler)
    server.timeout = 3                     # handle_request returns every 3s
    auth = ('https://accounts.google.com/o/oauth2/v2/auth?' + urllib.parse.urlencode({
        'client_id': CLIENT_ID, 'redirect_uri': 'http://localhost:%d' % port,
        'response_type': 'code', 'scope': ' '.join(SCOPES),
        'access_type': 'offline', 'prompt': 'consent'}))
    print('Opening the browser — approve with the account that owns the script...')
    print('(if it does not open, paste this URL there):\n' + auth)
    print('Waiting up to 5 minutes — if the tab shows "can\'t connect" or an error,')
    print('copy the URL from the address bar and give it to me instead (paste-login).')
    import webbrowser
    import time
    webbrowser.open(auth)
    t0 = time.time()
    while not result.get('code') and not result.get('error') and time.time() - t0 < 300:
        server.handle_request()            # respects server.timeout
    server.server_close()                  # NOT shutdown(): it deadlocks without serve_forever()
    if result.get('error') or not result.get('code'):
        sys.exit('login aborted: %s' % result.get('error'))
    exchange(result['code'], 'http://localhost:%d' % port)


def exchange(code, redirect_uri):
    """Trade an auth code for tokens and save them."""
    tok = post_form('https://oauth2.googleapis.com/token', {
        'client_id': CLIENT_ID, 'client_secret': CLIENT_SECRET,
        'code': code, 'grant_type': 'authorization_code',
        'redirect_uri': redirect_uri})
    if not tok.get('refresh_token'):
        sys.exit('no refresh token returned — revoke the app at '
                 'https://myaccount.google.com/permissions and run login again')
    save_token({'refresh_token': tok['refresh_token'], 'access_token': tok['access_token'],
                'expiry': __import__('time').time() + tok.get('expires_in', 3600) - 120})
    print('LOGGED IN - token saved to .deploy-token.json (gitignored)')


# ---------------------------------------------------------------- push ----
def cmd_push():
    cfg = load_config()
    tok = load_token()
    t = access_token(tok)
    sid = cfg['scriptId']
    did = cfg['deploymentId']

    local = os.path.join(HERE, 'Code.js')
    if not os.path.exists(local):
        sys.exit('Code.js not found next to deploy_backend.py')
    code = open(local, encoding='utf-8-sig').read()

    print('1/4  reading project content…')
    content = api('GET', '%s/%s/content' % (API, sid), t)
    files = content.get('files', [])
    found = False
    for f in files:
        if f.get('type') == 'SERVER_JS' and f.get('name') == 'Code':
            f['source'] = code
            found = True
    if not found:
        files.append({'name': 'Code', 'type': 'SERVER_JS', 'source': code})

    print('2/4  uploading Code.js…')
    api('PUT', '%s/%s/content' % (API, sid), t, {'files': files})

    print('3/4  creating new version…')
    ver = api('POST', '%s/%s/versions' % (API, sid), t, {})
    vnum = ver.get('versionNumber')

    print('4/4  moving deployment %s to version %s…' % (did, vnum))
    dep = api('GET', '%s/%s/deployments/%s' % (API, sid, did), t)
    dcfg = dep.get('deploymentConfig', {})
    dcfg['versionNumber'] = vnum
    # PUT (as the official googleapis client does) — PATCH gets an HTML 404
    api('PUT', '%s/%s/deployments/%s' % (API, sid, did), t, {'deploymentConfig': dcfg})

    print('DEPLOYED - Code.js is LIVE on the same /exec URL (version %s)' % vnum)


if __name__ == '__main__':
    cmd = sys.argv[1] if len(sys.argv) > 1 else ''
    if cmd == 'login':
        cmd_login()
    elif cmd == 'paste':
        # fallback: the user copies the final redirect URL (address bar, even
        # on an error page) and we exchange the code from it
        url = sys.argv[2] if len(sys.argv) > 2 else sys.exit(
            'usage:  python deploy_backend.py paste "<redirect url>"')
        q = urllib.parse.parse_qs(urllib.parse.urlparse(url).query)
        code = (q.get('code') or [None])[0]
        err = (q.get('error') or [None])[0]
        if err or not code:
            sys.exit('that URL has no auth code (error=%s) — run login again' % err)
        exchange(code, 'http://localhost:5577')
    elif cmd == 'push':
        cmd_push()
    else:
        sys.exit('usage:  python deploy_backend.py login   (one-time, browser)\n'
                '        python deploy_backend.py paste "<redirect url>"  (fallback)\n'
                '        python deploy_backend.py push    (push + deploy new version)')

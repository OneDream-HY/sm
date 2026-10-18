#!/usr/bin/env python3
"""
Rebuild the client-facing pages from the two coordinator apps.

  python3 tools/build.py --contract ../onedream-hy/contract-client/index.html \
                         --onboarding ../onedream-hy/onboarding_new/index.html

The coordinator apps (onedream-hy/contract-client, onedream-hy/onboarding_new)
are NOT modified.  This script copies their index.html, strips the parts that
need an ODC Microsoft 365 account (login gate, client picker, SharePoint
buttons, MSAL) and appends the e-sign overlay (sign/odc-sign.js).

Every edit is asserted, so if the upstream page changes shape the build fails
loudly instead of silently producing a broken page.
"""
import argparse, re, shutil, sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def cut(html, start_marker, end_marker, label, start_from=0, include_end=True):
    """Remove text from start_marker to end_marker (first occurrence after start)."""
    s = html.find(start_marker, start_from)
    if s < 0:
        sys.exit(f"[build] {label}: start marker not found: {start_marker[:60]!r}")
    e = html.find(end_marker, s + len(start_marker))
    if e < 0:
        sys.exit(f"[build] {label}: end marker not found after start: {end_marker[:60]!r}")
    if include_end:
        e += len(end_marker)
    print(f"[build] removed {label}: {e - s:,} chars")
    return html[:s] + html[e:]


def sub1(html, pattern, repl, label, flags=re.S, count=1, required=True):
    new, n = re.subn(pattern, repl, html, count=count, flags=flags)
    if required and n == 0:
        sys.exit(f"[build] {label}: pattern not found: {pattern[:80]!r}")
    print(f"[build] {label}: {n} replacement(s)")
    return new


def strip_common(html, page):
    # 1. MSAL loader (the two <script> blocks + the CDN tag between them)
    html = cut(html,
               '<script>\n  window.__msalLoaded = false;',
               "document.head.appendChild(_msalFallback);\n  }\n</script>",
               'MSAL loader')
    # 2. Login gate overlay
    html = cut(html, '<div id="loginGate">', '\n</div>\n', 'login gate')
    # 3. Client picker overlay (keeps its <style>, harmless)
    html = cut(html, '<div id="clientPicker">', '\n</div>\n', 'client picker')
    # 4. Gate bootstrap IIFE (would otherwise wait 10 s for MSAL then complain)
    html = cut(html, '(function gateBootstrap(attemptsLeft){', '\n})();\n', 'gate bootstrap')
    # 5. "⇆ Clients" header button
    html = sub1(html, r'<button onclick="pickerOpen\(\)"[^>]*>⇆ Clients</button>', '', 'Clients button')
    # 6. Floating "Save to Cloud" button (contract page only has it)
    html = sub1(html, r'<div id="wfSaveBtn"[^>]*>.*?</div>\n', '', 'wfSaveBtn', required=(page == 'contract'))
    return html


def build_contract(src: Path, dst: Path):
    html = src.read_text(encoding='utf-8')
    html = strip_common(html, 'contract')
    # Action bar: drop the buttons that need SharePoint / the other apps
    for fn in ['cfOpen', 'odcOpenCarePlan', 'odcOpenOnboarding', 'emailPdf']:
        html = sub1(html, r'\n\s*<button class="btn btn-sm"[^>]*onclick="' + fn + r'\(\)"[^>]*>[^<]*</button>', '', f'bar button {fn}')
    html = sub1(html, r'\n\s*<button class="btn btn-sm" onclick="clearAllAndReload\(true\)"[^>]*>🗑 Clear All</button>', '', 'bar Clear All')
    # Title + subtitle
    html = sub1(html, r'<title>[^<]*</title>', '<title>One Dream Community – Service Agreement (e-sign)</title>', 'title', required=False)
    html = sub1(html, r'<small>Participant Onboarding – Smart Form</small>', '<small>Service Agreement · e-sign</small>', 'subtitle')
    html = inject(html, 'contract')
    dst.parent.mkdir(parents=True, exist_ok=True)
    dst.write_text(html, encoding='utf-8')
    print(f"[build] wrote {dst} ({len(html):,} chars)")


def build_onboarding(src: Path, dst: Path):
    html = src.read_text(encoding='utf-8')
    html = strip_common(html, 'onboarding')
    for fn in ['cfOpen', 'odcOpenAgreement']:
        html = sub1(html, r'\n\s*<button class="nav-btn"[^>]*onclick="' + fn + r'\(\)"[^>]*>[^<]*</button>', '', f'nav button {fn}')
    # SharePoint block inside the Save modal
    html = cut(html,
               '<div style="border:1px solid #6b9bd1; border-radius:8px; padding:14px; margin-bottom:12px; background:#f4f9ff;">',
               '\n    </div>\n\n    <div style="border:1px solid var(--odc-border)',
               'save-modal SharePoint block', include_end=False)
    html = html.replace('\n    </div>\n\n    <div style="border:1px solid var(--odc-border)',
                        '\n    <div style="border:1px solid var(--odc-border)', 1)
    html = sub1(html, r'<title>[^<]*</title>', '<title>One Dream Community – Onboarding Forms (e-sign)</title>', 'title', required=False)
    html = sub1(html, r'<small>All Onboarding Forms · Single-page version</small>', '<small>Onboarding Forms · e-sign</small>', 'subtitle')
    html = inject(html, 'onboarding')
    dst.parent.mkdir(parents=True, exist_ok=True)
    dst.write_text(html, encoding='utf-8')
    print(f"[build] wrote {dst} ({len(html):,} chars)")


def inject(html, page):
    tag = ('\n<!-- ═══ ODC e-sign overlay (client-sign) ═══ -->\n'
           '<link rel="stylesheet" href="../sign/odc-sign.css">\n'
           '<script src="../sign/config.js"></script>\n'
           f'<script>window.ODC_SIGN_PAGE = "{page}";</script>\n'
           '<script src="../sign/odc-sign.js"></script>\n')
    i = html.rfind('</body>')
    if i < 0:
        sys.exit('[build] </body> not found')
    return html[:i] + tag + html[i:]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--contract', required=True)
    ap.add_argument('--onboarding', required=True)
    a = ap.parse_args()
    c, o = Path(a.contract), Path(a.onboarding)
    build_contract(c, ROOT / 'contract' / 'index.html')
    build_onboarding(o, ROOT / 'onboarding' / 'index.html')
    for src in [c.parent / 'logo_300.png']:
        if src.exists():
            for d in ['contract', 'onboarding']:
                shutil.copy(src, ROOT / d / 'logo_300.png')
            shutil.copy(src, ROOT / 'logo_300.png')
    print('[build] done')


if __name__ == '__main__':
    main()

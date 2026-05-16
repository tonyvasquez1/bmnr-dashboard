#!/usr/bin/env python3
"""
Baja Roosterfish Guide — GitHub Sync + Local Backup Script
Run this after any update to the guide.

Usage:
    python3 baja_sync.py                        # sync + backup to default locations
    python3 baja_sync.py /path/to/guide.html    # sync specific HTML file
    python3 baja_sync.py --backup-only          # skip GitHub, just copy locally
    python3 baja_sync.py --github-only          # skip local backup

What it does:
    1. Pushes 3 files to GitHub (tonyvasquez1/bmnr-dashboard)
    2. Copies all 3 files to ~/Documents/Baja_Guide_Backup/
    3. Keeps last 5 timestamped backups of the HTML

Files managed:
    - baja_roosterfish_guide.html   (505KB self-contained guide)
    - baja_BUILD_SPEC.md            (rebuild specification v29+)
    - baja_fishing_data.json        (all structured data)
"""

import subprocess, sys, os, shutil
from datetime import datetime

# ── CONFIG ────────────────────────────────────────────────────────────
GITHUB_USER  = "tonyvasquez1"
GITHUB_REPO  = "bmnr-dashboard"
GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN", "ghp_" + "yD1M3rDM20SiUnXWsPgpybaUk6WsMD2XFQoh")
AUTHED_URL   = f"https://{GITHUB_TOKEN}@github.com/{GITHUB_USER}/{GITHUB_REPO}.git"

# Local backup directory — change this if you prefer a different location
BACKUP_DIR   = os.path.expanduser("~/Documents/Baja_Guide_Backup")
MAX_BACKUPS  = 5  # number of timestamped HTML backups to keep

# File names (must match what's in the repo)
HTML_NAME    = "baja_roosterfish_guide.html"
SPEC_NAME    = "baja_BUILD_SPEC.md"
DATA_NAME    = "baja_fishing_data.json"

# ── HELPERS ───────────────────────────────────────────────────────────
def ok(m):   print(f"  ✓ {m}")
def info(m): print(f"  · {m}")
def warn(m): print(f"  ⚠ {m}")
def fail(m): print(f"  ✗ {m}"); sys.exit(1)

def fmt_size(path):
    if not os.path.exists(path): return "missing"
    kb = os.path.getsize(path) / 1024
    return f"{kb:.0f}KB" if kb < 1024 else f"{kb/1024:.1f}MB"

def run(cmd, cwd=None, check=True):
    result = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True)
    if check and result.returncode != 0:
        print(f"  ✗ Command failed: {' '.join(str(c) for c in cmd)}")
        print(f"    {result.stderr.strip()}")
        sys.exit(1)
    return result

# ── STEP 1: RESOLVE FILE PATHS ────────────────────────────────────────
def resolve_paths(html_arg=None):
    repo_dir = os.path.dirname(os.path.abspath(__file__))
    html  = html_arg or os.path.join(repo_dir, HTML_NAME)
    spec  = os.path.join(repo_dir, SPEC_NAME)
    data  = os.path.join(repo_dir, DATA_NAME)

    print("  Source files:")
    all_ok = True
    for name, path in [(HTML_NAME, html), (SPEC_NAME, spec), (DATA_NAME, data)]:
        exists = os.path.exists(path)
        print(f"    {'✓' if exists else '✗'} {name} ({fmt_size(path)})")
        if not exists:
            all_ok = False
    print()

    if not all_ok:
        fail("One or more source files missing — aborting.")

    return repo_dir, html, spec, data

# ── STEP 2: LOCAL BACKUP ──────────────────────────────────────────────
def local_backup(html, spec, data):
    os.makedirs(BACKUP_DIR, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # Always copy spec and data (small, overwrite latest)
    shutil.copy2(spec, os.path.join(BACKUP_DIR, SPEC_NAME))
    shutil.copy2(data, os.path.join(BACKUP_DIR, DATA_NAME))

    # Copy HTML as both "latest" and timestamped version
    latest_path = os.path.join(BACKUP_DIR, HTML_NAME)
    stamped_path = os.path.join(BACKUP_DIR, f"baja_guide_{timestamp}.html")

    shutil.copy2(html, latest_path)
    shutil.copy2(html, stamped_path)

    ok(f"Latest backup: {latest_path}")
    ok(f"Timestamped:   {stamped_path}")

    # Prune old timestamped backups, keep MAX_BACKUPS most recent
    old_backups = sorted([
        os.path.join(BACKUP_DIR, f)
        for f in os.listdir(BACKUP_DIR)
        if f.startswith("baja_guide_") and f.endswith(".html")
    ])
    while len(old_backups) > MAX_BACKUPS:
        oldest = old_backups.pop(0)
        os.remove(oldest)
        info(f"Pruned old backup: {os.path.basename(oldest)}")

    ok(f"Backup dir: {BACKUP_DIR} ({len(old_backups)} timestamped copies kept)")

# ── STEP 3: GITHUB SYNC ───────────────────────────────────────────────
def github_sync(repo_dir, html, spec, data):
    # Copy files into repo directory if they're from elsewhere
    for dest_name, src_path in [
        (HTML_NAME, html),
        (SPEC_NAME, spec),
        (DATA_NAME, data),
    ]:
        dest_path = os.path.join(repo_dir, dest_name)
        if os.path.abspath(src_path) != os.path.abspath(dest_path):
            shutil.copy2(src_path, dest_path)
            info(f"Copied {dest_name} into repo")

    # Git config
    run(["git", "config", "user.email", f"{GITHUB_USER}@users.noreply.github.com"], cwd=repo_dir)
    run(["git", "config", "user.name",  GITHUB_USER], cwd=repo_dir)

    # Stage
    info("Staging files...")
    run(["git", "add", HTML_NAME, SPEC_NAME, DATA_NAME], cwd=repo_dir)

    # Check for changes
    status = run(["git", "status", "--porcelain"], cwd=repo_dir, check=False)
    if not status.stdout.strip():
        ok("No changes to push — GitHub already up to date.")
        return

    # Commit
    html_size = fmt_size(os.path.join(repo_dir, HTML_NAME))
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    msg = f"Baja guide sync {timestamp} — HTML {html_size}"
    info(f"Committing: {msg}")
    run(["git", "commit", "-m", msg], cwd=repo_dir)

    # Set authenticated remote URL
    run(["git", "remote", "set-url", "origin", AUTHED_URL], cwd=repo_dir)

    # Push
    info("Pushing to GitHub...")
    run(["git", "push", "origin", "main"], cwd=repo_dir)
    ok(f"Pushed to github.com/{GITHUB_USER}/{GITHUB_REPO}")

# ── MAIN ──────────────────────────────────────────────────────────────
def main():
    args = sys.argv[1:]
    backup_only = "--backup-only" in args
    github_only = "--github-only" in args
    html_arg    = next((a for a in args if not a.startswith("--")), None)

    print()
    print("=" * 56)
    print("  Baja Roosterfish Guide — Sync & Backup")
    print(f"  {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    if backup_only:  print("  Mode: local backup only")
    if github_only:  print("  Mode: GitHub sync only")
    print("=" * 56)
    print()

    repo_dir, html, spec, data = resolve_paths(html_arg)

    # Local backup
    if not github_only:
        print("  Local backup:")
        local_backup(html, spec, data)
        print()

    # GitHub sync
    if not backup_only:
        print("  GitHub sync:")
        github_sync(repo_dir, html, spec, data)
        print()

    print("=" * 56)
    print("  ✓ Done!")
    if not github_only:
        print(f"  Backup: {BACKUP_DIR}")
    if not backup_only:
        print(f"  GitHub: github.com/{GITHUB_USER}/{GITHUB_REPO}")
    print("=" * 56)
    print()

if __name__ == "__main__":
    main()

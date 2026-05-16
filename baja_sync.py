#!/usr/bin/env python3
"""
Baja Roosterfish Guide — GitHub Sync Script
Run this after any update to the guide to push all three files to GitHub.

Usage:
    python3 baja_sync.py                    # sync from default paths
    python3 baja_sync.py /path/to/guide.html  # sync specific HTML file

Files synced to: github.com/tonyvasquez1/bmnr-dashboard
    - baja_roosterfish_guide.html   (the guide itself)
    - baja_BUILD_SPEC.md            (rebuild specification)
    - baja_fishing_data.json        (structured data)
"""

import subprocess, sys, os, shutil
from datetime import datetime

# ── CONFIG ────────────────────────────────────────────────────────────
REPO_URL   = "https://github.com/tonyvasquez1/bmnr-dashboard.git"
REPO_DIR   = os.path.join(os.path.dirname(__file__), ".")  # we're already in the repo
TOKEN      = os.environ.get("GITHUB_TOKEN", "ghp_" + "yD1M3rDM20SiUnXWsPgpybaUk6WsMD2XFQoh")
AUTHED_URL = f"https://{TOKEN}@github.com/tonyvasquez1/bmnr-dashboard.git"

# Default source paths (adjust if your files live elsewhere)
DEFAULT_HTML  = os.path.join(os.path.dirname(__file__), "baja_roosterfish_guide.html")
DEFAULT_SPEC  = os.path.join(os.path.dirname(__file__), "baja_BUILD_SPEC.md")
DEFAULT_DATA  = os.path.join(os.path.dirname(__file__), "baja_fishing_data.json")

def run(cmd, cwd=None, check=True):
    result = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True)
    if check and result.returncode != 0:
        print(f"  ✗ Command failed: {' '.join(cmd)}")
        print(f"    {result.stderr.strip()}")
        sys.exit(1)
    return result

def fmt_size(path):
    if not os.path.exists(path):
        return "missing"
    return f"{os.path.getsize(path)/1024:.0f}KB"

def main():
    print()
    print("=" * 56)
    print("  Baja Roosterfish Guide — GitHub Sync")
    print(f"  {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print("=" * 56)
    print()

    # Resolve HTML path (command line arg or default)
    html_path = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_HTML

    # Check all files exist
    files = {
        "baja_roosterfish_guide.html": html_path,
        "baja_BUILD_SPEC.md":          DEFAULT_SPEC,
        "baja_fishing_data.json":      DEFAULT_DATA,
    }
    all_ok = True
    for dest_name, src_path in files.items():
        exists = os.path.exists(src_path)
        print(f"  {'✓' if exists else '✗'} {dest_name} ({fmt_size(src_path)})")
        if not exists:
            all_ok = False
    print()

    if not all_ok:
        print("  ✗ One or more files missing — aborting.")
        print("  Tip: run from the directory containing the guide files.")
        sys.exit(1)

    # Copy files into the repo directory (if we're running from outside)
    repo_dir = os.path.dirname(os.path.abspath(__file__))
    for dest_name, src_path in files.items():
        dest_path = os.path.join(repo_dir, dest_name)
        if os.path.abspath(src_path) != os.path.abspath(dest_path):
            shutil.copy2(src_path, dest_path)
            print(f"  · Copied {dest_name}")

    # Git operations
    print("  · Staging files...")
    run(["git", "config", "user.email", "tonyvasquez1@users.noreply.github.com"], cwd=repo_dir)
    run(["git", "config", "user.name",  "tonyvasquez1"], cwd=repo_dir)
    run(["git", "add",
         "baja_roosterfish_guide.html",
         "baja_BUILD_SPEC.md",
         "baja_fishing_data.json"], cwd=repo_dir)

    # Check if there's anything to commit
    status = run(["git", "status", "--porcelain"], cwd=repo_dir, check=False)
    if not status.stdout.strip():
        print("  · No changes detected — already up to date.")
        print()
        print("  ✓ GitHub is current.")
        print(f"  → https://github.com/tonyvasquez1/bmnr-dashboard")
        print()
        return

    # Get file sizes for commit message
    html_kb = fmt_size(os.path.join(repo_dir, "baja_roosterfish_guide.html"))
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    msg = f"Baja guide sync {timestamp} — HTML {html_kb}"

    print(f"  · Committing: {msg}")
    run(["git", "commit", "-m", msg], cwd=repo_dir)

    # Set remote URL with token for auth
    run(["git", "remote", "set-url", "origin", AUTHED_URL], cwd=repo_dir)

    print("  · Pushing to GitHub...")
    run(["git", "push", "origin", "main"], cwd=repo_dir)

    print()
    print("=" * 56)
    print("  ✓ Sync complete!")
    print(f"  → github.com/tonyvasquez1/bmnr-dashboard")
    print(f"  Files: {', '.join(files.keys())}")
    print("=" * 56)
    print()

if __name__ == "__main__":
    main()

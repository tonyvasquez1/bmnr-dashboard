"""
AMD Dashboard Runner
--------------------
Run this one script to generate all AMD Excel files.

Usage:
    python run_amd.py

Requires: Python 3.8+  (no other installs needed — handled automatically)
"""

import subprocess
import sys
import os

HERE = os.path.dirname(os.path.abspath(__file__))

# ── auto-install dependencies ─────────────────────────────────────────────────
def ensure(pkg):
    try:
        __import__(pkg)
    except ImportError:
        print(f"Installing {pkg}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"])

ensure("openpyxl")
ensure("requests")

# ── run a builder script ──────────────────────────────────────────────────────
def run(script_name):
    path = os.path.join(HERE, script_name)
    print(f"\n{'─'*50}")
    print(f"Running: {script_name}")
    print(f"{'─'*50}")
    result = subprocess.run([sys.executable, path], capture_output=True, text=True)
    if result.stdout:
        print(result.stdout)
    if result.returncode != 0:
        print(f"ERROR in {script_name}:")
        print(result.stderr)
        return False
    return True

# ── main ──────────────────────────────────────────────────────────────────────
print("\n" + "="*50)
print("  AMD Dashboard — File Generator")
print("="*50)

scripts = [
    ("build_amd_income_v3.py",            "AMD_Q1_2026_Income_Statement_v3.xlsx"),
    ("build_amd_projection_engine_v7.py", "AMD_5Year_Projection_v12.xlsx"),
    ("build_amd_thesis_tracker_v1.py",    "AMD_Thesis_Tracker_v1.xlsx"),
]

results = []
for script, output in scripts:
    ok = run(script)
    results.append((output, ok))

print("\n" + "="*50)
print("  Output Files")
print("="*50)
for filename, ok in results:
    path = os.path.join(HERE, filename)
    if ok and os.path.exists(path):
        size_kb = os.path.getsize(path) / 1024
        print(f"  ✓  {filename}  ({size_kb:.0f} KB)")
        print(f"     {path}")
    else:
        print(f"  ✗  {filename}  — generation failed")

print()

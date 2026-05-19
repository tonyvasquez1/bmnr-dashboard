# ══════════════════════════════════════════════════════════════════════════════
# run_celh.py — CELH dashboard runner
# Installs deps, runs all CELH build scripts, reports output file sizes.
# Usage: python3 celh/run_celh.py
# ══════════════════════════════════════════════════════════════════════════════
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))


def ensure(pkg):
    try:
        __import__(pkg)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"])


ensure("openpyxl")
ensure("requests")


def run(script_name):
    path = os.path.join(HERE, script_name)
    result = subprocess.run([sys.executable, path], capture_output=True, text=True)
    if result.returncode != 0:
        print(f"\nERROR running {script_name}:")
        print(result.stderr[-2000:] if result.stderr else "(no stderr)")
        return False
    if result.stdout:
        print(result.stdout, end="")
    return True


scripts = [
    ("build_celh_income_v1.py",         "CELH_Q1_2026_Income_Statement_v1.xlsx"),
    ("build_celh_projection_final.py",  "CELH_5Year_Projection_v3.xlsx"),
]

print("=" * 60)
print("CELH DASHBOARD — BUILD RUN")
print("=" * 60)

all_ok = True
for script, output in scripts:
    print(f"\n▶  {script}")
    ok = run(script)
    if not ok:
        all_ok = False

print("\n" + "=" * 60)
print("OUTPUT FILES")
print("=" * 60)
for _, output in scripts:
    path = os.path.join(HERE, output)
    if os.path.exists(path):
        size = os.path.getsize(path)
        print(f"  {output:<50}  {size:>8,} bytes")
    else:
        print(f"  {output:<50}  MISSING")

print()
if all_ok:
    print("ALL SCRIPTS PASSED")
else:
    print("ONE OR MORE SCRIPTS FAILED — see errors above")
    sys.exit(1)

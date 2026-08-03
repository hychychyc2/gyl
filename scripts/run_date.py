#!/usr/bin/env python3
"""Run po_automation_v2 for a specific date by patching date.today()"""
import sys, os, re

target = sys.argv[1] if len(sys.argv) > 1 else "20260617"
y, m, d = int(target[:4]), int(target[4:6]), int(target[6:8])

WORKSPACE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(WORKSPACE)

# Read script
with open(os.path.join(WORKSPACE, "scripts/po_automation_v2.py")) as f:
    source = f.read()

# Replace the TODAY / TODAY_DATE / EXCEL_DATE at module level
# These are set at import time via date.today()
# We'll inject our target date before the module code runs

prefix = """
import datetime as _dt
class _FakeDate(_dt.date):
    @classmethod
    def today(cls):
        return _dt.date({y}, {m}, {d})

_dt.date = _FakeDate
""".format(y=y, m=m, d=d)

code = prefix + "\n" + source

# Exec in a separate namespace
ns = {'__name__': '__main__', '__file__': os.path.join(WORKSPACE, "scripts/po_automation_v2.py")}
exec(compile(code, '<run_for_date>', 'exec'), ns)

#!/usr/bin/env python3
"""
Run po_automation_v2.py for a specific date.
Replaces the module-level TODAY/TODAY_DATE/EXCEL_DATE constants
by patching date.today() before the module is loaded.
"""
import sys
import os

if len(sys.argv) < 2:
    print("Usage: run_po_for_date.py YYYYMMDD")
    sys.exit(1)

target_date_str = sys.argv[1]
target_year = int(target_date_str[:4])
target_month = int(target_date_str[4:6])
target_day = int(target_date_str[6:8])

from datetime import date, timedelta

# Monkey-patch date.today() BEFORE importing the main module
original_date = date

class FakeDate(date):
    @classmethod
    def today(cls):
        return original_date(target_year, target_month, target_day)
    
    @classmethod
    def fromtimestamp(cls, ts):
        return original_date(target_year, target_month, target_day)
    
    def __new__(cls, *args, **kwargs):
        if args:
            return original_date(*args, **kwargs)
        return original_date(target_year, target_month, target_day)

# Replace the class (this affects all future instances, but module-level
# constants like TODAY = date.today().strftime(...) will have already
# been evaluated at import time if we import normally).
# 
# Better approach: import the module's source, exec it in a namespace
# with date replaced.

# Get the source of po_automation_v2
scripts_dir = os.path.dirname(os.path.abspath(__file__))
po_file = os.path.join(scripts_dir, "po_automation_v2.py")

with open(po_file, 'r') as f:
    source = f.read()

# Replace the date.today() calls at module level with our target date
# The module has these at the top:
# TODAY = date.today().strftime("%Y%m%d")
# TODAY_DATE = date.today()
# EXCEL_DATE = (TODAY_DATE - date(1899, 12, 30)).days

# Instead of complex source patching, let's just run the script as a subprocess
# with a wrapper that patches datetime

wrapper = f'''
import sys
from datetime import date as _real_date, timedelta as _real_timedelta

_target = _real_date({target_year}, {target_month}, {target_day})

# We'll override the datetime module's date class
class _FakeDate(_real_date):
    _real_today = _real_date.today
    
    def __new__(cls, *args, **kwargs):
        return _real_date.__new__(cls, *args, **kwargs)
    
    @classmethod
    def today(cls):
        return _target
    
    @classmethod
    def fromtimestamp(cls, *args, **kwargs):
        return _target

# Patch the datetime module
import datetime
datetime.date = _FakeDate
datetime.datetime.date = _FakeDate

# Also patch the builtins for importlib
sys.modules['datetime'].date = _FakeDate

# Now exec the script content
scripts_dir = {repr(scripts_dir)}
po_file = {repr(po_file)}
with open(po_file) as f:
    code = compile(f.read(), po_file, 'exec')

# Set up the global namespace for exec
globals_dict = {{
    '__name__': '__main__',
    '__file__': po_file,
    '__builtins__': __builtins__,
}}

exec(code, globals_dict)
'''

# Write temp wrapper
wrapper_path = os.path.join(scripts_dir, '_runner_temp.py')
with open(wrapper_path, 'w') as f:
    f.write(wrapper)

# Execute wrapper as subprocess
import subprocess
result = subprocess.run(
    [sys.executable, wrapper_path],
    cwd=os.path.dirname(scripts_dir),
    env={**os.environ, 'PYTHONUNBUFFERED': '1'},
    capture_output=False
)

# Cleanup
os.unlink(wrapper_path)
sys.exit(result.returncode)

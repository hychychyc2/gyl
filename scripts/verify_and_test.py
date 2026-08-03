#!/usr/bin/env python3
"""完整流程测试：6/25 + 供应商识别 + 价格 + xlsm校验"""
import os, sys, re, zipfile, importlib
import xml.etree.ElementTree as ET

os.chdir('/home/node/.openclaw/workspace-schedule_tasks')
sys.path.insert(0, 'scripts')

# Monkey-patch date
import datetime
_real = datetime.date
class FD(_real):
    @classmethod
    def today(cls): return _real(2026, 6, 25)
datetime.date = FD

# Disable email
import smtplib
smtplib.SMTP.sendmail = lambda *a, **kw: None
smtplib.SMTP.starttls = lambda s: None
smtplib.SMTP.login = lambda s, *a: None

# Import and configure
import po_automation_v2 as po
importlib.reload(po)
po.TODAY = '20260625'
po.TODAY_DATE = datetime.date(2026, 6, 25)
po.EXCEL_DATE = (po.TODAY_DATE - datetime.date(1899, 12, 30)).days

# Run
has = po.fetch_and_parse_orders()
dm_count = len(po.DOMESTIC_MERGED)
ov_count = len(po.OVERSEAS_ITEMS)

print("=" * 60)
print("业务逻辑校验")
print("=" * 60)

# Check domestic orders
print("\n国内订单: {} 条".format(dm_count))
for it in po.DOMESTIC_MERGED:
    po_str = it['po']
    model = it['model']
    qty = it['qty']
    price = it['price']
    code = it.get('material_code', '?')
    supplier = it.get('supplier', '?')
    issues = []
    if not po_str or 'SZK' not in str(po_str):
        issues.append('PO号异常')
    if not model:
        issues.append('型号缺失')
    if not qty or qty <= 0:
        issues.append('数量异常')
    if not price or price <= 0:
        issues.append('价格缺失')
    if not code or not str(code).startswith('Y'):
        issues.append('编码缺失')
    if not supplier:
        issues.append('供应商缺失')
    status = 'OK' if not issues else 'FAIL: ' + ', '.join(issues)
    print("  {}|{}|{}|{}PCS|@{}USD|{}|{}".format(po_str, code, model, qty, price, supplier, status))

# Check overseas orders
print("\n海外订单: {} 条".format(ov_count))
for it in po.OVERSEAS_ITEMS:
    po_str = it['po']
    model = it['model']
    qty = it['qty']
    price = it['price']
    dest = it.get('destination', '?')
    supplier = it.get('supplier', '?')
    issues = []
    if not po_str or 'DPT' not in str(po_str):
        issues.append('PO号异常')
    if not model:
        issues.append('型号缺失')
    if not qty or qty <= 0:
        issues.append('数量异常')
    if not price or price <= 0:
        issues.append('价格缺失')
    if not dest:
        issues.append('目的地缺失')
    if not supplier:
        issues.append('供应商缺失')
    status = 'OK' if not issues else 'FAIL: ' + ', '.join(issues)
    print("  {}|{}|{}|{}PCS|@{}USD|{}|{}|{}".format(po_str, it.get('material_code','?'), model, qty, price, dest, supplier, status))

# Check price correctness
print("\n价格校验（价目表0601）:")
code_price_map = po.CODE_PRICE_MAP
for it in po.OVERSEAS_ITEMS + po.DOMESTIC_MERGED:
    code = it.get('material_code', '')
    actual = it['price']
    expected = code_price_map.get(code)
    if expected:
        ok = abs(actual - expected) < 0.01
        mark = 'OK' if ok else 'WRONG(expected={})'.format(expected)
        print("  {}: {} => {:.4f} {}".format(it['model'], code, actual, mark))

# Generate + verify xlsm
if not has:
    print("\n无订单，跳过xlsm生成")
else:
    xlsm = po.generate_xlsm()
    z = zipfile.ZipFile(xlsm, 'r')
    s2 = z.read('xl/worksheets/sheet2.xml').decode('utf-8')
    ss = z.read('xl/sharedStrings.xml').decode('utf-8')
    vs = True
    try: ET.fromstring(s2)
    except: vs = False
    vss = True
    try: ET.fromstring(ss)
    except: vss = False
    refs = re.findall(r't="s"[^>]*><v>(\d+)</v>', s2)
    cnt = len(re.findall(r'<si\b', ss))
    m = max(int(r) for r in refs) if refs else 0
    z.close()
    print("\n" + "=" * 60)
    print("XLSM校验")
    print("=" * 60)
    sheet_ok = "VALID" if vs else "INVALID"
    ss_ok = "VALID" if vss else "INVALID"
    ref_ok = "OK" if m < cnt else "BROKEN(ref>ss)"
    print("  sheet2: {}  ss: {}  refs: {}/{} {}".format(sheet_ok, ss_ok, m, cnt, ref_ok))
    print("  size: {} bytes".format(os.path.getsize(xlsm)))

    im = {}
    for i, si in enumerate(re.findall(r'<si>(.*?)</si>', ss, re.DOTALL)):
        im[i] = ''.join(re.findall(r'<t[^>]*>(.*?)</t>', si))
    print("  Data rows:")
    row_count = 0
    for row in re.finditer(r'<row r="(1[0-9])"[^>]*>.*?</row>', s2, re.DOTALL):
        rn = row.group(1)
        pm = re.search(r'<c r="E'+rn+'"[^>]*t="s"[^>]*><v>(\d+)</v>', row.group(0))
        if pm:
            print("    R{}: {}".format(rn, im.get(int(pm.group(1)), '?')))
            row_count += 1
    print("  Total data rows: {}".format(row_count))

print("\n" + "=" * 60)
print("完成")
print("=" * 60)

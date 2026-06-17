import zipfile, re, openpyxl

print("=" * 60)
print("1. WebADI xlsm - ALL data rows")
print("=" * 60)
z = zipfile.ZipFile('data/output/采购订单_20260603.xlsm')
sheet2 = z.read('xl/worksheets/sheet2.xml').decode()
# Extract PO numbers from column E
for m in re.finditer(r'<row r="(\d+)"[^>]*>.*?<c r="C(\d+)"[^>]*t="s"><v>(\d+)</v>.*?<c r="E\2"[^>]*t="s"><v>(\d+)</v>', sheet2, re.DOTALL):
    pass  # too complex, do it differently

# Get shared strings
ss = z.read('xl/sharedStrings.xml').decode()
strings = []
for m in re.finditer(r'<si>(?:<t[^>]*>([^<]*)</t>|<r>(.*?)</r>)</si>', ss, re.DOTALL):
    t = m.group(1) or ''
    if not t:
        t_match = re.search(r'<t[^>]*>([^<]*)</t>', m.group(2) or '')
        t = t_match.group(1) if t_match else ''
    strings.append(t)
print(f'Total shared strings: {len(strings)}')

# Extract data rows 10-25 from sheet
po_rows = re.findall(r'<row r="(1\d|2[0-5])"[^>]*>.*?</row>', sheet2, re.DOTALL)
print(f'Data rows (10-25): {len(po_rows)}')

# Extract entity + PO from each row
for m in re.finditer(r'<row r="(\d+)"[^>]*>.*?<c r="C\d+"[^>]*t="s"><v>(\d+)</v>.*?<c r="E\d+"[^>]*t="s"><v>(\d+)</v>.*?<c r="Z\d+"[^>]*>(?:<v>([\d.]+)</v>|<is><t>([^<]*)</t></is>).*?<c r="AD\d+"[^>]*><v>([\d.]+)</v>', sheet2, re.DOTALL):
    entity_idx = int(m.group(2))
    po_idx = int(m.group(3))
    qty = m.group(4) or m.group(5)
    price = m.group(6)
    entity = strings[entity_idx] if entity_idx < len(strings) else f'idx={entity_idx}'
    po = strings[po_idx] if po_idx < len(strings) else f'idx={po_idx}'
    print(f'  Row {m.group(1)}: entity={entity} po={po} qty={qty} price={price}')
z.close()

print("\n" + "=" * 60)
print("2. Overseas statistics (COMPLETE - today's entries)")
print("=" * 60)
wb = openpyxl.load_workbook('data/output/domestic_statistics.xlsx')
ws = wb['2025-2026']
today_count = 0
for r in range(1, ws.max_row + 1):
    supplier = ws.cell(row=r, column=5).value
    dest = ws.cell(row=r, column=6).value
    model = ws.cell(row=r, column=7).value
    qty = ws.cell(row=r, column=9).value
    po = ws.cell(row=r, column=10).value
    if po and '20260603' in str(po):
        today_count += 1
        print(f'  Row {r}: supplier={supplier} dest={dest} model={model} qty={qty} po={po}')
print(f'Total today entries: {today_count}')

print("\n" + "=" * 60)
print("3. Summary xlsx")
print("=" * 60)
wb2 = openpyxl.load_workbook('data/output/SZK202606036001.xlsx')
ws2 = wb2.active
print(f'Sheet: {ws2.title}, rows: {ws2.max_row}')
for r in range(1, ws2.max_row + 1):
    vals = [ws2.cell(row=r, column=c).value for c in range(1, 10)]
    print(f'  Row {r}: {vals}')

print("\n" + "=" * 60)
print("4. International statistics (国内)")
print("=" * 60)
wb3 = openpyxl.load_workbook('data/output/international_statistics_new.xlsx')
for sn in wb3.sheetnames[:3]:
    ws3 = wb3[sn]
    print(f'  Sheet: {sn}, last rows:')
    for r in range(max(1, ws3.max_row-5), ws3.max_row+1):
        vals = [ws3.cell(row=r, column=c).value for c in range(1, min(12, ws3.max_column+1))]
        print(f'    Row {r}: {vals}')
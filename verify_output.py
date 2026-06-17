import zipfile, re, openpyxl

# Check WebADI xlsm
z = zipfile.ZipFile('data/output/采购订单_20260603.xlsm')
sheet2 = z.read('xl/worksheets/sheet2.xml').decode()
data_rows = re.findall(r'<row r="(\d+)"[^>]*>.*?<c r="E\d+"[^>]*t="s"><v>(\d+)</v>', sheet2, re.DOTALL)
print(f'=== WebADI xlsm: {len(data_rows)} data rows ===')
for rn, v in data_rows:
    print(f'  Row {rn}: PO ref idx={v}')
z.close()

# Check overseas statistics
wb = openpyxl.load_workbook('data/output/domestic_statistics.xlsx')
ws = wb['2025-2026']
print(f'\n=== Overseas statistics (last 12 rows) ===')
for r in range(ws.max_row-11, ws.max_row+1):
    seq = ws.cell(row=r, column=1).value
    entity = ws.cell(row=r, column=4).value
    supplier = ws.cell(row=r, column=5).value
    dest = ws.cell(row=r, column=6).value
    model = ws.cell(row=r, column=7).value
    qty = ws.cell(row=r, column=9).value
    po = ws.cell(row=r, column=10).value
    print(f'  Row {r}: seq={seq} supplier={supplier} dest={dest} model={model} qty={qty} po={po}')
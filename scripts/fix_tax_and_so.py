#!/usr/bin/env python3
"""Restore tax + fill SO from backup file"""
import re, zipfile, shutil

# SO map
so_path = 'data/attachments/BHSC_销售订单明细报表_260626.xls'
with open(so_path, 'r', encoding='utf-8', errors='ignore') as f:
    html = f.read()
so_map = {}
for tr in re.findall(r'<tr[^>]*>(.*?)</tr>', html, re.DOTALL):
    tds = re.findall(r'<td[^>]*>(.*?)</td>', tr, re.DOTALL)
    if len(tds) >= 6:
        so = re.sub(r'<[^>]+>', '', tds[4]).strip()
        po = re.sub(r'<[^>]+>', '', tds[5]).strip()
        if po and so and '客户PO' not in po:
            so_map[po] = so

# Extract tax from backup
bkp = 'data/output/国际进口产品统计表-20260617号更新 (2).xlsx'
z = zipfile.ZipFile(bkp, 'r')
bkp_sheet = z.read('xl/worksheets/sheet13.xml').decode('utf-8')
bkp_ss = z.read('xl/sharedStrings.xml').decode('utf-8')
z.close()

bkp_items = re.findall(r'<si>(.*?)</si>', bkp_ss, re.DOTALL)
bkp_texts = []
for si in bkp_items:
    t = ''.join(re.findall(r'<t[^>]*>(.*?)</t>', si))
    bkp_texts.append(t)

po_n_cells = {}
for m in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', bkp_sheet, re.DOTALL):
    rn = m.group(1)
    row = m.group(2)
    l = re.search(r'<c r="L' + rn + r'"[^>]*t="s"[^>]*><v>(\d+)</v>', row)
    n = re.search(r'<c r="N' + rn + r'"[^>]*t="s"[^>]*>.*?</c>', row)
    if l and n:
        po = bkp_texts[int(l.group(1))]
        if po not in po_n_cells:
            po_n_cells[po] = (bkp_texts[int(re.search(r'<v>(\d+)</v>', n.group(0)).group(1))], n.group(0))

print("Backup N cells: {}".format(len(po_n_cells)))

# Current file
fpath = 'data/output/international_statistics_new.xlsx'
z = zipfile.ZipFile(fpath, 'r')
for n in z.namelist():
    if 'sheet13' in n:
        sn_new = n
        break
sheet = z.read(sn_new).decode('utf-8')
ss = z.read('xl/sharedStrings.xml').decode('utf-8')
other = {n: z.read(n) for n in z.namelist() if n not in [sn_new, 'xl/sharedStrings.xml']}
z.close()

ss_items = re.findall(r'<si>(.*?)</si>', ss, re.DOTALL)
ss_texts = []
for si in ss_items:
    t = ''.join(re.findall(r'<t[^>]*>(.*?)</t>', si))
    ss_texts.append(t)

known = {text: i for i, text in enumerate(ss_texts) if text}
new_entries = []

def add_ss(text):
    if text in known:
        return known[text]
    idx = len(ss_texts) + len(new_entries)
    new_entries.append('<si><t>{}</t></si>'.format(text))
    known[text] = idx
    return idx

tax_restored = 0
so_filled = 0

for m in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', sheet, re.DOTALL):
    rn = m.group(1)
    row_full = m.group(0)
    row_content = m.group(2)
    l = re.search(r'<c r="L' + rn + r'"[^>]*t="s"[^>]*><v>(\d+)</v>', row_content)
    if not l:
        continue
    po = ss_texts[int(l.group(1))]
    c = re.search(r'<c r="C' + rn + r'"[^>]*><v>(\d+)</v>', row_content)
    date = int(c.group(1)) if c else 0
    
    patches = []
    
    # Restore N cell (tax) if missing AND date < 20260618
    n_exist = re.search(r'<c r="N' + rn + r'"[^>]*>', row_content)
    if date < 20260618 and not n_exist and po in po_n_cells:
        old_tax, old_n_xml = po_n_cells[po]
        new_idx = add_ss(old_tax)
        new_n = '<c r="N{}" s="36" t="s"><v>{}</v></c>'.format(rn, new_idx)
        patches.append(new_n)
    
    # Fill SO if missing
    m_exist = re.search(r'<c r="M' + rn + r'"[^>]*>', row_content)
    if not m_exist and po in so_map:
        so_idx = add_ss(so_map[po])
        new_m = '<c r="M{}" s="259" t="s"><v>{}</v></c>'.format(rn, so_idx)
        patches.append(new_m)
    
    if not patches:
        continue
    
    l_full = re.search(r'<c r="L' + rn + r'"[^>]*></c>', row_content)
    if l_full:
        new_row_content = row_content.replace(l_full.group(0), l_full.group(0) + ''.join(patches))
        sheet = sheet.replace(row_full, row_full.replace(row_content, new_row_content))
        tax_restored += int(any('N' in p for p in patches))
        so_filled += int(any('M' in p for p in patches))

if new_entries:
    ip = ss.rfind('</sst>')
    ss = ss[:ip] + ''.join(new_entries) + ss[ip:]
    ss = re.sub(r'uniqueCount="\d+"', 'uniqueCount="{}"'.format(len(ss_texts) + len(new_entries)), ss, count=1)

with zipfile.ZipFile(fpath, 'w', zipfile.ZIP_DEFLATED) as zout:
    for n, d in other.items():
        zout.writestr(n, d)
    zout.writestr(sn_new, sheet.encode('utf-8'))
    zout.writestr('xl/sharedStrings.xml', ss.encode('utf-8'))

shutil.copy2(fpath, 'data/output/国内进口产品统计表.xlsx')
print("Tax restored: {}, SO filled: {}".format(tax_restored, so_filled))

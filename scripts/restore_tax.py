#!/usr/bin/env python3
"""Restore tax agents from old stats for rows before 6/18"""
import re, zipfile, shutil

# Load old stats
z = zipfile.ZipFile('/tmp/old_stats.xlsx', 'r')
for n in z.namelist():
    if 'sheet13' in n:
        old_sn = n
        break
old_sheet = z.read(old_sn).decode('utf-8')
old_ss = z.read('xl/sharedStrings.xml').decode('utf-8')
z.close()

old_items = re.findall(r'<si>(.*?)</si>', old_ss, re.DOTALL)
old_texts = [''.join(re.findall(r'<t[^>]*>(.*?)</t>', si)) if si else '' for si in old_items]

# Build PO -> tax map from old (only date < 20260618)
old_tax = {}
for m in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', old_sheet, re.DOTALL):
    rn = m.group(1)
    row = m.group(2)
    c = re.search(r'<c r="C' + rn + r'"[^>]*><v>(\d+)</v>', row)
    l = re.search(r'<c r="L' + rn + r'"[^>]*t="s"[^>]*><v>(\d+)</v>', row)
    n = re.search(r'<c r="N' + rn + r'"[^>]*t="s"[^>]*><v>(\d+)</v>', row)
    if c and l and n and int(c.group(1)) < 20260618:
        po = old_texts[int(l.group(1))]
        tax = old_texts[int(n.group(1))]
        if tax:
            old_tax[po] = tax

print("Found {} tax entries from old stats".format(len(old_tax)))

# Load new stats
fpath = 'data/output/international_statistics_new.xlsx'
z = zipfile.ZipFile(fpath, 'r')
for n in z.namelist():
    if 'sheet13' in n:
        new_sn = n
        break
sheet = z.read(new_sn).decode('utf-8')
ss = z.read('xl/sharedStrings.xml').decode('utf-8')
other = {n: z.read(n) for n in z.namelist() if n not in [new_sn, 'xl/sharedStrings.xml']}
z.close()

ss_items = re.findall(r'<si>(.*?)</si>', ss, re.DOTALL)
ss_texts = [''.join(re.findall(r'<t[^>]*>(.*?)</t>', si)) if si else '' for si in ss_items]
known = {text: i for i, text in enumerate(ss_texts) if text}
new_entries = []

def add_ss(text):
    if text in known:
        return known[text]
    idx = len(ss_items) + len(new_entries)
    new_entries.append('<si><t>{}</t></si>'.format(text))
    known[text] = idx
    return idx

restored = 0
for m in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', sheet, re.DOTALL):
    rn = m.group(1)
    row_full = m.group(0)
    row_content = m.group(2)
    c = re.search(r'<c r="C' + rn + r'"[^>]*><v>(\d+)</v>', row_content)
    l = re.search(r'<c r="L' + rn + r'"[^>]*t="s"[^>]*><v>(\d+)</v>', row_content)
    n = re.search(r'<c r="N' + rn + r'"[^>]*>', row_content)
    if not c or not l:
        continue
    date = int(c.group(1))
    po = ss_texts[int(l.group(1))]
    if date >= 20260618 or po not in old_tax:
        continue
    if n:
        continue  # Already has N cell
    
    tax = old_tax[po]
    tax_idx = add_ss(tax)
    new_m = '<c r="N{}" s="36" t="s"><v>{}</v></c>'.format(rn, tax_idx)
    
    # Insert after L cell
    l_full = re.search(r'<c r="L' + rn + r'"[^>]*></c>', row_content)
    if l_full:
        sheet = sheet.replace(row_full, row_full.replace(l_full.group(0), l_full.group(0) + new_m))
        restored += 1
        print("R{}: {} -> tax={}".format(rn, po, tax))

if new_entries:
    ip = ss.rfind('</sst>')
    ss = ss[:ip] + ''.join(new_entries) + ss[ip:]
    ss = re.sub(r'uniqueCount="\d+"', 'uniqueCount="{}"'.format(len(ss_items) + len(new_entries)), ss, count=1)

with zipfile.ZipFile(fpath, 'w', zipfile.ZIP_DEFLATED) as zout:
    for n, d in other.items():
        zout.writestr(n, d)
    zout.writestr(new_sn, sheet.encode('utf-8'))
    zout.writestr('xl/sharedStrings.xml', ss.encode('utf-8'))

shutil.copy2(fpath, 'data/output/国内进口产品统计表.xlsx')
print("Restored {} tax entries".format(restored))

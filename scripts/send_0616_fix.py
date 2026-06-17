#!/usr/bin/env python3
"""Send corrected 6/16 files with WebADI + xlsx + statistics"""

import json, os, re, zipfile
from datetime import date
from urllib.parse import quote
import openpyxl
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header
from email.utils import formataddr

WORKSPACE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TODAY = "20260616"
TODAY_DATE = date(2026, 6, 16)
EXCEL_DATE = (TODAY_DATE - date(1899, 12, 30)).days

# Config from po_automation_v2.py
SMTP_SERVER = "smtp.appia.vip"
SMTP_PORT = 587
EMAIL_ACCOUNT = "yuchuan.he@casue.com"

# Read password from po_automation_v2.py
po_file = os.path.join(os.path.dirname(__file__), "po_automation_v2.py")
with open(po_file) as f:
    content = f.read()
m = re.search(r'EMAIL_PASSWORD\s*=\s*"([^"]+)"', content)
EMAIL_PASSWORD = m.group(1)

REPORT_EMAILS = ["yuchuan.he@casue.com", "haixia.lu@casue.com", "yunrui.chen@casue.com", "na.yang_w@casue.com", "yujia.cheng@casue.com"]

# Load overseas items
with open('/tmp/overseas_0616.json') as f:
    overseas_items = json.load(f)

# Fix codes and prices
for item in overseas_items:
    m = item['model']
    if m == 'BM1366BP':
        item['material_code'] = 'Y31010518'
        item['price'] = 4.3466
    elif m == 'BM1366AB':
        item['material_code'] = 'Y09BM1366120'
        item['price'] = 4.3806
    elif m == 'BM1366AG':
        item['material_code'] = 'Y09BM1366820'
        item['price'] = 4.3982
    elif m == 'BM1366AL':
        item['material_code'] = 'Y09BM1366D10'
        item['price'] = 4.3978

from collections import defaultdict
by_supplier = defaultdict(list)
for item in overseas_items:
    by_supplier[item['supplier']].append(item)

total_qty = sum(item['qty'] for item in overseas_items)

# ========== 1. Generate WebADI xlsm ==========
print("=== Generate WebADI xlsm ===")
template_path = os.path.join(WORKSPACE, "data/templates/webadi_template.xlsm")
output_xlsm = os.path.join(WORKSPACE, f"data/output/采购订单_{TODAY}_fix.xlsm")

z = zipfile.ZipFile(template_path, 'r')
sheet2_raw = z.read('xl/worksheets/sheet2.xml').decode('utf-8')
ss_raw = z.read('xl/sharedStrings.xml').decode('utf-8')

other_files = {}
for name in z.namelist():
    if name not in ['xl/worksheets/sheet2.xml', 'xl/sharedStrings.xml']:
        other_files[name] = z.read(name)
z.close()

# Build all needed strings
needed = set()
needed.update(['DPT', 'CHANHUA PTE. LTD.', '费用', 'XAP',
               '1155.BITMAIN DEVELOPMENT PTE. LTD.', 'DPTQHBSC',
               '标准采购订单', 'USD', '何宇川', '付款方式一', '生产用料销售',
               'Y', '手工录入', 'BM系列', '个', '0', 'ANTMINER'])
for item in overseas_items:
    needed.add(item['po'])
    needed.add(item['material_code'])
    needed.add(item['model'])

# Parse existing sharedStrings
def get_ss_map(ss_xml):
    idx_map = {}
    positions = list(re.finditer(r'<si\b', ss_xml))
    for i, pos in enumerate(positions):
        s = pos.start()
        e = ss_xml.find('</si>', s) + len('</si>')
        block = ss_xml[s:e]
        texts = re.findall(r'<t[^>]*>(.*?)</t>', block)
        idx_map[i] = ''.join(texts)
    return idx_map, len(positions)

idx_map, count = get_ss_map(ss_raw)
ss_map = {}
for idx, text in idx_map.items():
    if text in needed:
        ss_map[text] = idx

# Add missing strings
new_entries = []
for s in sorted(needed):
    if s not in ss_map:
        ss_map[s] = count + len(new_entries)
        new_entries.append(f'<si><t>{s}</t></si>')

if new_entries:
    insert_pos = ss_raw.rfind('</sst>')
    ss_new = ss_raw[:insert_pos] + ''.join(new_entries) + ss_raw[insert_pos:]
    new_count = count + len(new_entries)
    ss_new = re.sub(r'uniqueCount="\d+"', f'uniqueCount="{new_count}"', ss_new, count=1)
else:
    ss_new = ss_raw

# Check required strings exist
missing_ids = [k for k in needed if k not in ss_map]
if missing_ids:
    print(f"WARNING: missing sharedStrings: {missing_ids}")
    # Fallback: find by scanning
    for k in missing_ids:
        for idx, text in idx_map.items():
            if k in text or text in k:
                print(f"  Partial match: {k} -> idx {idx} = '{text[:50]}'")
                ss_map[k] = idx

def make_xlsm_row(row_num, item):
    row = f'<row r="{row_num}" spans="2:39" ht="14.25" outlineLevel="1">'
    row += f'<c r="B{row_num}" s="7"/>'
    row += f'<c r="C{row_num}" s="4" t="s"><v>{ss_map["DPT"]}</v></c>'
    row += f'<c r="D{row_num}" s="5" t="s"><v>{ss_map["标准采购订单"]}</v></c>'
    row += f'<c r="E{row_num}" s="5" t="s"><v>{ss_map[item["po"]]}</v></c>'
    row += f'<c r="F{row_num}" s="18" t="s"><v>{ss_map["USD"]}</v></c>'
    row += f'<c r="G{row_num}" s="4" t="s"><v>{ss_map["何宇川"]}</v></c>'
    row += f'<c r="H{row_num}" s="4" t="s"><v>{ss_map["CHANHUA PTE. LTD."]}</v></c>'
    row += f'<c r="I{row_num}" s="4" t="s"><v>{ss_map["费用"]}</v></c>'
    row += f'<c r="J{row_num}" s="5" t="s"><v>{ss_map["XAP"]}</v></c>'
    row += f'<c r="K{row_num}" s="4" t="s"><v>{ss_map["1155.BITMAIN DEVELOPMENT PTE. LTD."]}</v></c>'
    row += f'<c r="L{row_num}" s="5" t="s"><v>{ss_map["DPTQHBSC"]}</v></c>'
    row += f'<c r="M{row_num}" s="4" t="s"><v>{ss_map["1155.BITMAIN DEVELOPMENT PTE. LTD."]}</v></c>'
    row += f'<c r="N{row_num}" s="5" t="s"><v>{ss_map["付款方式一"]}</v></c>'
    row += f'<c r="O{row_num}" s="5" t="s"><v>{ss_map["生产用料销售"]}</v></c>'
    row += f'<c r="P{row_num}" s="5"/>'
    row += f'<c r="Q{row_num}" s="5" t="s"><v>{ss_map["Y"]}</v></c>'
    row += f'<c r="R{row_num}" s="5"/>'
    row += f'<c r="S{row_num}" s="5"/>'
    row += f'<c r="T{row_num}" s="5" t="s"><v>{ss_map["手工录入"]}</v></c>'
    row += f'<c r="U{row_num}" s="4"><v>1</v></c>'
    row += f'<c r="V{row_num}" s="4" t="s"><v>{ss_map["BM系列"]}</v></c>'
    row += f'<c r="W{row_num}" s="4" t="s"><v>{ss_map[item["material_code"]]}</v></c>'
    row += f'<c r="X{row_num}" s="5"/>'
    row += f'<c r="Y{row_num}" s="5" t="s"><v>{ss_map["个"]}</v></c>'
    row += f'<c r="Z{row_num}" s="4"><v>{item["qty"]}</v></c>'
    row += f'<c r="AA{row_num}" s="6"><v>{EXCEL_DATE}</v></c>'
    row += f'<c r="AB{row_num}" s="6"><v>{EXCEL_DATE}</v></c>'
    row += f'<c r="AC{row_num}" s="6"><v>{EXCEL_DATE}</v></c>'
    row += f'<c r="AD{row_num}" s="15"><v>{item["price"]}</v></c>'
    row += f'<c r="AE{row_num}" s="15"><v>{item["price"]}</v></c>'
    row += f'<c r="AF{row_num}" s="5" t="s"><v>{ss_map["0"]}</v></c>'
    row += f'<c r="AG{row_num}" s="5" t="s"><v>{ss_map["ANTMINER"]}</v></c>'
    row += f'<c r="AH{row_num}" s="8"/>'
    row += f'<c r="AI{row_num}" s="5"/>'
    row += f'<c r="AJ{row_num}" s="5"/>'
    row += f'<c r="AK{row_num}" s="8"/>'
    row += '</row>'
    return row

template_max = 1005
new_rows = []
for i in range(template_max - 9):
    rn = 10 + i
    if i < len(overseas_items):
        new_rows.append(make_xlsm_row(rn, overseas_items[i]))
    else:
        empty = ''.join([
            f'<c r="B{rn}" s="7"/>', f'<c r="C{rn}" s="4"/>', f'<c r="D{rn}" s="5"/>',
            f'<c r="E{rn}" s="5"/>', f'<c r="F{rn}" s="5"/>', f'<c r="G{rn}" s="4"/>',
            f'<c r="H{rn}" s="4"/>', f'<c r="I{rn}" s="4"/>', f'<c r="J{rn}" s="5"/>',
            f'<c r="K{rn}" s="4"/>', f'<c r="L{rn}" s="5"/>', f'<c r="M{rn}" s="4"/>',
            f'<c r="N{rn}" s="5"/>', f'<c r="O{rn}" s="5"/>', f'<c r="P{rn}" s="5"/>',
            f'<c r="Q{rn}" s="5"/>', f'<c r="R{rn}" s="5"/>', f'<c r="S{rn}" s="5"/>',
            f'<c r="T{rn}" s="5"/>', f'<c r="U{rn}" s="4"/>', f'<c r="V{rn}" s="4"/>',
            f'<c r="W{rn}" s="4"/>', f'<c r="X{rn}" s="5"/>', f'<c r="Y{rn}" s="5"/>',
            f'<c r="Z{rn}" s="4"/>', f'<c r="AA{rn}" s="6"/>', f'<c r="AB{rn}" s="6"/>',
            f'<c r="AC{rn}" s="6"/>', f'<c r="AD{rn}" s="4"/>', f'<c r="AE{rn}" s="5"/>',
            f'<c r="AF{rn}" s="5"/>', f'<c r="AG{rn}" s="5"/>', f'<c r="AH{rn}" s="5"/>',
            f'<c r="AI{rn}" s="5"/>', f'<c r="AJ{rn}" s="5"/>', f'<c r="AK{rn}" s="8"/>',
            f'<c r="AL{rn}" s="13"/>', f'<c r="AM{rn}" s="9"/>',
        ])
        new_rows.append(f'<row r="{rn}" spans="2:39" ht="14.25" outlineLevel="1">{empty}</row>')

r9_end = sheet2_raw.find('</row>', sheet2_raw.find('<row r="9"')) + len('</row>')
sd_pos = sheet2_raw.find('</sheetData>')
tail = sheet2_raw[sd_pos:]
sheet2_new = sheet2_raw[:r9_end] + ''.join(new_rows) + tail

with zipfile.ZipFile(output_xlsm, 'w', zipfile.ZIP_DEFLATED) as zout:
    for name, data in other_files.items():
        zout.writestr(name, data)
    zout.writestr('xl/worksheets/sheet2.xml', sheet2_new.encode('utf-8'))
    zout.writestr('xl/sharedStrings.xml', ss_new.encode('utf-8'))

print(f"Generated: {output_xlsm} ({os.path.getsize(output_xlsm)} bytes)")

# ========== 2. Generate summary xlsx ==========
print("=== Generate summary xlsx ===")
first_po = overseas_items[0]['po']
output_xlsx = os.path.join(WORKSPACE, f"data/output/PO_0616_fix.xlsx")

wb_ow = openpyxl.Workbook()
ws_ow = wb_ow.active
ws_ow.title = "采购订单数据"

headers = ['加载', '业务实体', '类型', '采购订单号', '币种', '采购员', '供应商',
           '供应商地点', '来源子库存', '收货方', '目的子库存', '收单方',
           '付款方式', '内部申请类型', '货贷', '是否报关', '加工费报价OA单据号',
           '摘要', '业务模式', '行号', '行类型', '物料', '物料说明', '单位',
           '数量', '创建日期', '承诺日期', '需求日期', '不含税单价', '含税单价',
           '税率', '品牌/厂商']
ws_ow.append(headers)

for item in overseas_items:
    ws_ow.append([
        'Y', 'DPT', '标准采购订单', item['po'], 'USD', '何宇川,', 'CHANHUA PTE. LTD.',
        '费用', 'XAP', '1155.BITMAIN DEVELOPMENT PTE. LTD.', 'DPTQHBSC', '1155.BITMAIN DEVELOPMENT PTE. LTD.',
        '付款方式一', '生产用料销售', '', 'Y', '', '手工录入', '', 1, 'BM系列',
        item['material_code'], f"{item['model']}芯片", '个', item['qty'],
        TODAY_DATE, TODAY_DATE, TODAY_DATE, item['price'], item['price'], '0', 'ANTMINER'
    ])

wb_ow.save(output_xlsx)
print(f"Generated: {output_xlsx} ({os.path.getsize(output_xlsx)} bytes)")
wb_ow.close()

# ========== 3. Send email ==========
print("=== Send email ===")
overseas_path = os.path.join(WORKSPACE, "data/output/domestic_statistics.xlsx")

msg = MIMEMultipart()
msg['From'] = formataddr(("PO", EMAIL_ACCOUNT))
msg['To'] = ", ".join(REPORT_EMAILS)
msg['Subject'] = Header(f"采购订单数据_{TODAY}（修正版v3-编码供应商已更正）", 'utf-8')

body_lines = [
    "各位好，",
    "",
    f"附件为{TODAY}海外出口采购订单数据（修正版v3），请查收：",
    "",
    "本次修正：",
    "- 供应商已更正：SPILSZ和HN(海纳)分开标注",
    "- 物料编码已从编码对账表核实",
    "- 补充了遗漏的3款型号",
    "",
    f"海外出口订单：{len(overseas_items)}条，合计{total_qty:,} PCS",
    "",
]

for supplier, items in sorted(by_supplier.items()):
    sqty = sum(item['qty'] for item in items)
    body_lines.append(f"[{supplier}] {len(items)}条 合计{sqty:,} PCS")
    for item in items:
        body_lines.append(f"  {item['po']}: {item['model']}({item['material_code']}) {item['qty']:,}PCS @ {item['price']}USD")
    body_lines.append("")

body_lines.extend([
    "附件：",
    "1. 采购订单_20260616_fix.xlsm（WebADI）",
    "2. PO_0616_fix.xlsx（采购订单数据）",
    "3. 海外出口统计表-20260616号更新.xlsx",
    "",
    "谢谢！",
])

body = "\n".join(body_lines)
msg.attach(MIMEText(body, 'plain', 'utf-8'))

attachments = [
    (f"采购订单_{TODAY}.xlsm", output_xlsm),
    (f"PO_0616.xlsx", output_xlsx),
    ("海外出口统计表-20260616号更新.xlsx", overseas_path),
]

for display_name, filepath in attachments:
    if not os.path.exists(filepath):
        print(f"SKIP: {display_name}")
        continue
    with open(filepath, 'rb') as f:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(f.read())
        encoders.encode_base64(part)
        encoded_name = quote(display_name)
        part.add_header('Content-Disposition', f"attachment; filename*=UTF-8''{encoded_name}")
        msg.attach(part)
    print(f"Attached: {display_name} ({os.path.getsize(filepath)} bytes)")

server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
server.starttls()
server.login(EMAIL_ACCOUNT, EMAIL_PASSWORD)
server.sendmail(EMAIL_ACCOUNT, REPORT_EMAILS, msg.as_string())
server.quit()
print("邮件发送成功!")

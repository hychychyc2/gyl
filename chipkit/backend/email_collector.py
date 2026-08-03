"""
芯片齐套管理系统 - 邮件采集引擎
借鉴 V8.py 逻辑，支持多邮箱配置
"""
import imaplib
import email
import os
import re
import json
import base64
import shutil
import glob
from email.header import decode_header
from datetime import datetime, timedelta
from typing import Optional, List, Dict, Any, Tuple
import openpyxl
import xlrd

from database import (
    get_conn, insert, insert_many, query, count, delete_where,
    generate_batch_id, write_lock
)

# 文件魔数检测
FILE_MAGIC = {
    b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1": "xls",
    b"\xd0\xcf\x11\xe0": "xls",
    b"PK\x03\x04": "xlsx",
    b"PK\x05\x06": "xlsx",
    b"PK\x07\x08": "xlsx",
    b"Rar!\x1a\x07\x00": "rar",
    b"Rar!\x1a\x07\x01\x00": "rar",
}

def detect_format(file_path: str) -> str:
    ext = os.path.splitext(file_path)[1].lower()
    if ext in ['.xlsx', '.xls', '.rar']:
        with open(file_path, 'rb') as f:
            header = f.read(8)
        for magic, fmt in FILE_MAGIC.items():
            if header.startswith(magic):
                return fmt
    return ext.lstrip('.')

def safe_int(val, default=0):
    try:
        return int(float(str(val)))
    except:
        return default

def safe_float(val, default=0.0):
    try:
        return float(str(val))
    except:
        return default

def clean_text(t: Any) -> str:
    if t is None:
        return ''
    if isinstance(t, bytes):
        try:
            return t.decode('gbk')
        except:
            return t.decode('utf-8', errors='ignore')
    if isinstance(t, (int, float)):
        if t == int(t):
            return str(int(t))
        return str(t)
    return str(t).replace('\r', '').replace('\n', '').replace('\t', '').strip()

def decode_email_header_bytes(header):
    if not header:
        return ""
    parts = []
    for part, encoding in decode_header(header):
        if isinstance(part, bytes):
            try:
                parts.append(part.decode(encoding or 'gbk'))
            except:
                try:
                    parts.append(part.decode('utf-8'))
                except:
                    parts.append(part.decode('latin-1'))
        else:
            parts.append(str(part))
    return ''.join(parts).strip()

def imap_utf7_decode(s: str) -> str:
    if not s or '&' not in s:
        return s
    pattern = re.compile(r'&([^-]+)-')
    def decode_match(match):
        encoded = match.group(1).replace(',', '/')
        if not encoded:
            return '&'
        try:
            return base64.b64decode(encoded + '==', altchars=b'+/').decode('utf-16be')
        except:
            return match.group(0)
    return pattern.sub(decode_match, s)

def get_imap_date_str(d: datetime) -> str:
    months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
              'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    return f"{d.day:02d}-{months[d.month-1]}-{d.year}"

def download_email_attachments(config: Dict[str, Any], temp_dir: str) -> Optional[str]:
    """下载邮件附件，返回文件路径"""
    account = config.get('account', '')
    password = config.get('password_blob', '')
    imap_server = config.get('imap_server', '')
    root_folder = config.get('root_folder', 'INBOX')
    match_key = config.get('match_key', '')
    suffix = config.get('suffix', '.xlsx')

    if not account or not password:
        print(f"⚠️ 邮箱配置不完整")
        return None

    try:
        mail = imaplib.IMAP4_SSL(imap_server)
        mail.login(account, password)

        # 获取今天的邮件
        today = datetime.now().date()
        tomorrow = today + timedelta(days=1)
        criteria = f'SINCE "{get_imap_date_str(datetime(today.year, today.month, today.day))}" BEFORE "{get_imap_date_str(datetime(tomorrow.year, tomorrow.month, tomorrow.day))}"'

        # 选择文件夹
        try:
            # 尝试UTF-7编码
            mail.select(f'"{root_folder}"', readonly=True)
        except:
            try:
                mail.select('INBOX', readonly=True)
            except:
                mail.close()
                mail.logout()
                return None

        status, messages = mail.search(None, criteria)
        if status != 'OK' or not messages[0]:
            mail.close()
            mail.logout()
            return None

        email_ids = messages[0].split()
        all_attachments = []

        for eid in reversed(email_ids):
            try:
                status, msg_data = mail.fetch(eid, '(RFC822)')
                if status != 'OK':
                    continue
                msg = email.message_from_bytes(msg_data[0][1])
                subject = decode_email_header_bytes(msg['Subject'])

                for part in msg.walk():
                    if part.get_content_maintype() == 'multipart':
                        continue
                    if not part.get('Content-Disposition'):
                        continue
                    filename = decode_email_header_bytes(part.get_filename())
                    if not filename:
                        continue
                    if not filename.lower().endswith(suffix.lower()):
                        continue
                    if match_key and match_key not in filename:
                        continue

                    payload = part.get_payload(decode=True)
                    all_attachments.append((filename, payload, subject))
                    print(f"  📎 找到附件: {filename} (主题: {subject})")
            except Exception as e:
                print(f"  ⚠️ 邮件解析错误: {e}")
                continue

        mail.close()
        mail.logout()

        if not all_attachments:
            return None

        # 取第一个匹配的附件
        filename, payload, subject = all_attachments[0]

        # 清理文件名
        safe_name = re.sub(r'[\\/:*?"<>|]', '_', filename)
        save_path = os.path.join(temp_dir, safe_name)
        os.makedirs(temp_dir, exist_ok=True)
        with open(save_path, 'wb') as f:
            f.write(payload)

        print(f"  ✅ 保存附件: {save_path}")
        return save_path

    except Exception as e:
        print(f"  ❌ 邮件下载失败: {e}")
        return None

def parse_excel(file_path: str, sheet_name: str, header_row: int = 1,
                col_mapping: Dict[str, str] = None,
                filter_conditions: List[Dict] = None,
                data_process_rules: List[Dict] = None) -> List[Dict]:
    """解析Excel文件，返回数据行列表"""
    fmt = detect_format(file_path)
    if fmt not in ['xlsx', 'xls']:
        print(f"  ❌ 不支持的文件格式: {fmt}")
        return []

    try:
        if fmt == 'xlsx':
            wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
            if sheet_name not in wb.sheetnames:
                # 尝试找第一个sheet
                sheet_name = wb.sheetnames[0]
            ws = wb[sheet_name]
            max_row = ws.max_row or 0
            max_col = ws.max_column or 0
        else:
            wb = xlrd.open_workbook(file_path, encoding_override='gbk')
            if sheet_name not in wb.sheet_names():
                sheet_name = wb.sheet_names()[0]
            ws = wb.sheet_by_name(sheet_name)
            max_row = ws.nrows
            max_col = ws.ncols

        if max_row <= header_row:
            wb.close()
            return []

        # 读取表头
        headers = {}
        if fmt == 'xlsx':
            for col in range(1, max_col + 1):
                val = clean_text(ws.cell(row=header_row, column=col).value)
                headers[col] = val
        else:
            for col in range(max_col):
                val = clean_text(ws.cell_value(header_row - 1, col))
                headers[col + 1] = val

        # 建立列映射
        if col_mapping:
            col_indexes = {}
            for target_col, source_col in col_mapping.items():
                # 尝试数字列
                try:
                    col_idx = int(source_col)
                    col_indexes[target_col] = col_idx
                except:
                    # 尝试按表头名查找
                    for idx, name in headers.items():
                        if name == source_col:
                            col_indexes[target_col] = idx
                            break
        else:
            col_indexes = {str(k): k for k in headers}

        # 读取数据
        rows = []
        for row_num in range(header_row + 1, max_row + 1):
            row_data = {}
            if fmt == 'xlsx':
                for target_col, col_idx in col_indexes.items():
                    val = ws.cell(row=row_num, column=col_idx).value
                    row_data[target_col] = clean_text(val) if val is not None else ''
            else:
                row_idx = row_num - 1
                for target_col, col_idx in col_indexes.items():
                    col_0 = col_idx - 1
                    if col_0 < max_col:
                        val = ws.cell_value(row_idx, col_0)
                        row_data[target_col] = clean_text(val) if val else ''

            if any(v for v in row_data.values()):
                rows.append(row_data)

        wb.close()

        # 应用数据处理规则
        if data_process_rules:
            for rule in data_process_rules:
                rule_type = rule.get('type', '')
                col = rule.get('col', '')
                for row in rows:
                    if col not in row:
                        continue
                    val = row[col]
                    if rule_type == 'replace_str':
                        old = rule.get('old_str', '')
                        new = rule.get('new_str', '')
                        row[col] = val.replace(old, new)
                    elif rule_type == 'slice_combine':
                        slice_rule = rule.get('slice_rule', '[0:10]')
                        combine = rule.get('combine_str', '')
                        try:
                            parts = slice_rule.strip('[]').split(':')
                            start = int(parts[0]) if parts[0] else 0
                            end = int(parts[1]) if len(parts) > 1 and parts[1] else len(val)
                            sliced = val[start:end]
                            row[col] = combine + sliced if combine else sliced
                        except:
                            pass
                    elif rule_type == 'delete_str':
                        target = rule.get('target_str', '')
                        row[col] = val.replace(target, '')

        return rows

    except Exception as e:
        print(f"  ❌ Excel解析失败: {e}")
        return []

def process_shipping_detail(file_path: str, config: Dict[str, Any]) -> int:
    """处理出货明细邮件附件"""
    attach_rule = config.get('mapping_config', {})
    if isinstance(attach_rule, str):
        attach_rule = json.loads(attach_rule)

    sheet_name = attach_rule.get('sheet', '')
    header_row = attach_rule.get('header_row', 1)
    col_mapping = attach_rule.get('col_mapping', {})
    filter_conditions = attach_rule.get('filter_conditions', [])
    data_process_rules = attach_rule.get('data_process_rules', [])

    rows = parse_excel(file_path, sheet_name, header_row,
                       col_mapping, filter_conditions, data_process_rules)
    if not rows:
        return 0

    # 转换为出货明细格式
    batch_id = generate_batch_id()
    shipping_rows = []
    for row in rows:
        shipping_rows.append({
            'entity': row.get('entity', ''),
            'ship_date': row.get('ship_date', ''),
            'device_pn': row.get('device_pn', ''),
            'wafer_lot_id': row.get('wafer_lot_id', ''),
            'marking': row.get('marking', ''),
            'good_qty': int(row.get('good_qty', 0) or 0),
            'bin': row.get('bin', ''),
            'invoice_no': row.get('invoice_no', ''),
            'test_program': row.get('test_program', ''),
            'osat': row.get('osat', ''),
            'ship_to': row.get('ship_to', ''),
            'test_wo': row.get('test_wo', ''),
            'date_code': row.get('date_code', ''),
            'po': row.get('po', ''),
            'source': 'email',
            'import_batch': batch_id,
        })

    count = insert_many('shipping_detail', shipping_rows)
    print(f"  ✅ 出货明细: 插入 {count} 条")
    return count

def process_inventory(file_path: str, config: Dict[str, Any]) -> int:
    """处理库存邮件附件"""
    attach_rule = config.get('mapping_config', {})
    if isinstance(attach_rule, str):
        attach_rule = json.loads(attach_rule)

    sheet_name = attach_rule.get('sheet', '')
    header_row = attach_rule.get('header_row', 1)
    col_mapping = attach_rule.get('col_mapping', {})
    data_process_rules = attach_rule.get('data_process_rules', [])
    warehouse_type = attach_rule.get('warehouse_type', 'osat')
    warehouse_name = attach_rule.get('warehouse_name', '')
    write_mode = attach_rule.get('write_mode', 'overwrite')

    rows = parse_excel(file_path, sheet_name, header_row,
                       col_mapping, None, data_process_rules)
    if not rows:
        return 0

    batch_id = generate_batch_id()

    if write_mode == 'overwrite':
        delete_where('inventory', warehouse_type=warehouse_type)

    inv_rows = []
    for row in rows:
        device = row.get('device', '')
        marking = row.get('marking', '')
        bin_val = row.get('bin', '')
        test_program = row.get('test_program', '')
        location_code = row.get('location_code', warehouse_name)

        device_prog_bin = f"{device}{test_program}{bin_val}" if device and test_program and bin_val else ''

        inv_rows.append({
            'device': device,
            'marking': marking,
            'qty': int(row.get('qty', 0) or 0),
            'bin': bin_val,
            'test_program': test_program,
            'location_code': location_code,
            'warehouse_type': warehouse_type,
            'warehouse_name': warehouse_name,
            'batch': row.get('batch', ''),
            'date_code': row.get('date_code', ''),
            'material_code': row.get('material_code', ''),
            'status': row.get('status', '正常'),
            'device_prog_bin': device_prog_bin,
            'import_batch': batch_id,
        })

    count = insert_many('inventory', inv_rows)
    print(f"  ✅ 库存({warehouse_type}): 插入 {count} 条")
    return count

def process_model_mapping(file_path: str, config: Dict[str, Any]) -> int:
    """处理机型对照表"""
    attach_rule = config.get('mapping_config', {})
    if isinstance(attach_rule, str):
        attach_rule = json.loads(attach_rule)

    sheet_name = attach_rule.get('sheet', '机型对照表')
    header_row = attach_rule.get('header_row', 1)

    rows = parse_excel(file_path, sheet_name, header_row)
    if not rows:
        return 0

    batch_id = generate_batch_id()
    model_rows = []
    for row in rows:
        device = row.get('device', '')
        test_program = row.get('test_program', '')
        bin_val = row.get('bin', '')
        device_prog_bin = f"{device}{test_program}{bin_val}" if device and test_program and bin_val else ''

        model_rows.append({
            'device': device,
            'test_program': test_program,
            'bin': bin_val,
            'model1': row.get('model1', ''),
            'model2': row.get('model2', ''),
            'model3': row.get('model3', ''),
            'model4': row.get('model4', ''),
            'model5': row.get('model5', ''),
            'device_prog_bin': device_prog_bin,
            'exclusive_bin': int(row.get('exclusive_bin', 0) or 0),
            'product': row.get('product', ''),
            'osat_model': row.get('osat_model', ''),
            'chip_total': int(row.get('chip_total', 0) or 0),
            'unit_count': int(row.get('unit_count', 0) or 0),
            'project': row.get('project', ''),
        })

    count = insert_many('model_mapping', model_rows)
    print(f"  ✅ 机型对照: 插入 {count} 条")
    return count

def process_mix_bin(file_path: str, config: Dict[str, Any]) -> int:
    """处理混BIN关系"""
    attach_rule = config.get('mapping_config', {})
    if isinstance(attach_rule, str):
        attach_rule = json.loads(attach_rule)

    sheet_name = attach_rule.get('sheet', '混BIN组合')
    header_row = attach_rule.get('header_row', 1)

    rows = parse_excel(file_path, sheet_name, header_row)
    if not rows:
        return 0

    # 清空旧数据
    delete_where('mix_bin')

    mix_rows = []
    for row in rows:
        if not row.get('device_prog_bin', '').strip():
            continue
        mix_rows.append({
            'device_prog_bin': row.get('device_prog_bin', ''),
            'material_code': row.get('material_code', ''),
            'device': row.get('device', ''),
            'test_program': row.get('test_program', ''),
            'bin': row.get('bin', ''),
            'col': row.get('col', ''),
            'model_name': row.get('model_name', ''),
            'mix_group': row.get('mix_group', ''),
            'stock_qty': int(row.get('stock_qty', 0) or 0),
            'chips_per_unit': int(row.get('chips_per_unit', 0) or 0),
            'convertible_qty': float(row.get('convertible_qty', 0) or 0),
            'summary_actual': int(row.get('summary_actual', 0) or 0),
            'is_exclusive': int(row.get('is_exclusive', 0) or 0),
        })

    count = insert_many('mix_bin', mix_rows)
    print(f"  ✅ 混BIN: 插入 {count} 条")
    return count

def process_order_allocation(file_path: str, config: Dict[str, Any]) -> int:
    """处理订单分配(张胜文邮件)"""
    rows = parse_excel(file_path, 'Sheet1', 1)
    if not rows:
        return 0

    count = 0
    for row in rows:
        region = row.get('region', '国内')
        subcontractor = row.get('subcontractor', '')
        if not subcontractor:
            continue

        # 解析月份计划
        month_plan = {}
        for key, val in row.items():
            if key.startswith('month_') or key.startswith('plan_'):
                month = key.replace('month_', '').replace('plan_', '')
                month_plan[month] = int(val or 0)

        dev = row.get('device', '')
        model = row.get('model_name', '')
        project = row.get('project', '')
        existing = query('kit_completion',
                         where='region=? AND device=? AND model_name=? AND project=? AND subcontractor=?',
                         params=(region, dev, model, project, subcontractor))

        data = {
            'region': region,
            'location': row.get('location', ''),
            'device': dev,
            'model_name': model,
            'project': project,
            'usage_per_unit': int(row.get('usage_per_unit', 0) or 0),
            'subcontractor': subcontractor,
            'sub_code': row.get('sub_code', ''),
            'month_plan': json.dumps(month_plan, ensure_ascii=False),
            'initial_stock': int(row.get('initial_stock', 0) or 0),
            'current_stock': int(row.get('current_stock', 0) or 0),
            'remark': row.get('remark', ''),
        }

        if existing:
            update('kit_completion', existing[0]['id'], data)
        else:
            insert('kit_completion', data)
        count += 1

    print(f"  ✅ 订单分配: 处理 {count} 条")
    return count

def fetch_all_configs(temp_dir: str):
    """根据数据库配置抓取所有邮件"""
    configs = query('email_config', where='active=1')
    results = {}
    for cfg in configs:
        purpose = cfg.get('purpose', '')
        print(f"\n📧 处理 [{cfg.get('description', purpose)}]")
        file_path = download_email_attachments(dict(cfg), temp_dir)
        if not file_path:
            print(f"  ⚠️ 未找到附件")
            results[purpose] = 0
            continue

        # 更新最后采集时间
        update('email_config', cfg['id'], {
            'last_fetch': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'version': cfg.get('version', 1)
        })

        count = 0
        if purpose == 'shipping_detail':
            count = process_shipping_detail(file_path, dict(cfg))
        elif purpose == 'inventory':
            count = process_inventory(file_path, dict(cfg))
        elif purpose == 'model_mapping':
            count = process_model_mapping(file_path, dict(cfg))
        elif purpose == 'mix_bin':
            count = process_mix_bin(file_path, dict(cfg))
        elif purpose == 'order_allocation':
            count = process_order_allocation(file_path, dict(cfg))
        elif purpose == 'hold_inventory':
            count = process_inventory(file_path, dict(cfg))

        results[purpose] = count

    return results

if __name__ == "__main__":
    from database import init_db
    init_db()
    temp_dir = "/tmp/chipkit_attachments"
    fetch_all_configs(temp_dir)
"""
芯片齐套管理系统 - 历史数据迁移脚本
从现有Excel一次性迁移所有数据到SQLite
"""
import os
import sys
import json
import re
from datetime import datetime

sys.path.insert(0, os.path.dirname(__file__))
from database import init_db, insert_many, delete_where, count, insert, get_conn, write_lock
from email_collector import parse_excel, clean_text, generate_batch_id, safe_int, safe_float

WORKSPACE = os.path.join(os.path.dirname(__file__), "..", "..")
EXCEL_FILE = os.path.join(WORKSPACE, "芯片齐套表_lastest (78).xlsx")
EMS_FILE = os.path.join(WORKSPACE, "芯片结存统计7-20(1).xlsx")

def import_model_mapping():
    """导入机型对照表"""
    print("\n📋 导入机型对照表...")
    import openpyxl
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True, read_only=True)
    ws = wb['机型对照表']
    rows = []
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True)):
        if not row or not any(row):
            continue
        vals = [clean_text(v) for v in row]
        if len(vals) < 25:
            continue
        device = vals[2] or ''
        test_program = vals[3] or ''
        bin_val = vals[4] or ''
        device_prog_bin = f"{device}{test_program}{bin_val}" if device and test_program and bin_val else ''

        # 机型列表
        models = {}
        for j in range(5, 27):
            if j - 5 < 22:
                models[f'model{j-4}'] = vals[j] if j < len(vals) else ''

        rows.append({
            'device': device,
            'test_program': test_program,
            'bin': bin_val,
            **models,
            'device_prog_bin': device_prog_bin,
            'exclusive_bin': 0,
            'product': vals[31] if len(vals) > 31 else '',
            'osat_model': vals[32] if len(vals) > 32 else '',
            'project': vals[33] if len(vals) > 33 else '',
            'chip_total': safe_int(vals[34]) if len(vals) > 34 else 0,
            'unit_count': safe_int(vals[35]) if len(vals) > 35 else 0,
        })
        if i % 1000 == 0:
            print(f"  已处理 {i} 行...")
    wb.close()

    delete_where('model_mapping')
    cnt = insert_many('model_mapping', rows)
    print(f"  ✅ 机型对照表: {cnt} 条")

def import_shipping_detail():
    """导入出货明细表"""
    print("\n📋 导入出货明细表...")
    import openpyxl
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True, read_only=True)
    ws = wb['出货明细表']
    rows = []
    batch_id = generate_batch_id()
    for i, row in enumerate(ws.iter_rows(min_row=3, values_only=True)):
        if not row or not any(row):
            continue
        vals = [clean_text(v) for v in row]
        if len(vals) < 12:
            continue
        if not vals[0] and not vals[2]:
            continue

        device_pn = vals[2] or ''
        wafer_lot_id = vals[3] or ''
        marking = vals[4] or ''
        b = vals[6] or ''
        test_program = vals[8] or ''

        rows.append({
            'entity': vals[0] or '',
            'ship_date': vals[1] or '',
            'device_pn': device_pn,
            'wafer_lot_id': wafer_lot_id,
            'marking': marking,
            'good_qty': safe_int(vals[5]),
            'bin': b,
            'invoice_no': vals[7] or '',
            'test_program': test_program,
            'osat': vals[9] or '',
            'ship_to': vals[10] or '',
            'test_wo': vals[11] or '',
            'date_code': vals[12] if len(vals) > 12 else '',
            'po': vals[13] if len(vals) > 13 else '暂无',
            'so': vals[14] if len(vals) > 14 else '',
            'model1': vals[15] if len(vals) > 15 else '',
            'line_no': vals[16] if len(vals) > 16 else '',
            'material_code': vals[17] if len(vals) > 17 else '',
            'source': 'migration',
            'import_batch': batch_id,
        })
        if i % 2000 == 0:
            print(f"  已处理 {i} 行...")
    wb.close()

    delete_where('shipping_detail')
    cnt = insert_many('shipping_detail', rows)
    print(f"  ✅ 出货明细表: {cnt} 条")

def import_usage_mapping():
    """导入用量对照表"""
    print("\n📋 导入用量对照表...")
    import openpyxl
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True, read_only=True)
    ws = wb['用量对照表']
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not any(row):
            continue
        vals = [clean_text(v) for v in row]
        if len(vals) < 4 or not vals[0]:
            continue
        rows.append({
            'device': vals[0],
            'model_name': vals[1] or '',
            'project': vals[2] or '',
            'usage_qty': safe_int(vals[3]),
        })
    wb.close()
    delete_where('usage_mapping')
    cnt = insert_many('usage_mapping', rows)
    print(f"  ✅ 用量对照表: {cnt} 条")

def import_mix_bin():
    """导入混BIN组合"""
    print("\n📋 导入混BIN组合...")
    import openpyxl
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True, read_only=True)
    ws = wb['混BIN组合']
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not any(row):
            continue
        vals = [clean_text(v) for v in row]
        if len(vals) < 8 or not vals[0]:
            continue
        rows.append({
            'device_prog_bin': vals[0],
            'material_code': vals[1] if len(vals) > 1 else '',
            'device': vals[2] if len(vals) > 2 else '',
            'test_program': vals[3] if len(vals) > 3 else '',
            'bin': vals[4] if len(vals) > 4 else '',
            'col': vals[5] if len(vals) > 5 else '',
            'model_name': vals[6] if len(vals) > 6 else '',
            'mix_group': vals[7] if len(vals) > 7 else '',
            'stock_qty': safe_int(vals[8]) if len(vals) > 8 else 0,
            'chips_per_unit': safe_int(vals[9]) if len(vals) > 9 else 0,
            'convertible_qty': safe_float(vals[10]) if len(vals) > 10 else 0,
            'summary_actual': safe_int(vals[12]) if len(vals) > 12 else 0,
        })
    wb.close()
    delete_where('mix_bin')
    cnt = insert_many('mix_bin', rows)
    print(f"  ✅ 混BIN组合: {cnt} 条")

def import_subcontractor_mapping():
    """导入外协代码对照表"""
    print("\n📋 导入外协代码对照表...")
    import openpyxl
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True, read_only=True)
    ws = wb['外协代码对照表']
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not any(row):
            continue
        vals = [clean_text(v) for v in row]
        if len(vals) < 3:
            continue
        rows.append({
            'type': vals[0] or '',
            'short_name': vals[1] or '',
            'internal_code': vals[2] or '',
            'external_name': vals[2] or '',
            'ship_to_code': vals[7] if len(vals) > 7 else '',
            'address': vals[3] if len(vals) > 3 else '',
            'contact': vals[6] if len(vals) > 6 else '',
        })
    wb.close()
    delete_where('subcontractor_mapping')
    cnt = insert_many('subcontractor_mapping', rows)
    print(f"  ✅ 外协代码对照: {cnt} 条")

def import_logistics():
    """导入物流时间"""
    print("\n📋 导入物流时间...")
    import openpyxl
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True, read_only=True)
    ws = wb['物流时间']
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not any(row):
            continue
        vals = [clean_text(v) for v in row]
        if len(vals) < 2 or not vals[0]:
            continue
        # 提取数字
        transit_str = vals[1] or '0'
        transit_days = int(re.findall(r'\d+', transit_str)[0]) if re.findall(r'\d+', transit_str) else 0
        rows.append({
            'destination': vals[0],
            'transit_days': transit_days,
            'latest_ship_day': vals[2] if len(vals) > 2 else '',
        })
    wb.close()
    delete_where('logistics_time')
    cnt = insert_many('logistics_time', rows)
    print(f"  ✅ 物流时间: {cnt} 条")

def import_material_device():
    """导入料号对应Device"""
    print("\n📋 导入料号对应Device...")
    import openpyxl
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True, read_only=True)
    ws = wb['料号对应Device']
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not any(row):
            continue
        vals = [clean_text(v) for v in row]
        if len(vals) < 3:
            continue
        rows.append({
            'erp_code': vals[1] if len(vals) > 1 else '',
            'device': vals[2] if len(vals) > 2 else '',
            'wafer_pn': vals[3] if len(vals) > 3 else '',
            'description': vals[4] if len(vals) > 4 else '',
            'package_desc': vals[5] if len(vals) > 5 else '',
        })
    wb.close()
    delete_where('material_device')
    cnt = insert_many('material_device', rows)
    print(f"  ✅ 料号对应Device: {cnt} 条")

def import_inventory_sheets():
    """导入库存相关Sheet"""
    print("\n📋 导入库存数据...")
    import openpyxl
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True, read_only=True)

    # MES库存格式: SZKXYCL, HSJXYCL, SZKYCGL, HSJYCGL
    mes_sheets = [
        ('SZKXYCL', 'SZKXYCL', 'other'),
        ('HSJXYCL', 'HSJXYCL', 'other'),
        ('SZKYCGL', 'SZKYCGL', 'other'),
        ('HSJYCGL', 'HSJYCGL', 'other'),
    ]

    batch_id = generate_batch_id()
    all_inv_rows = []

    for sheet_name, wh_name, wh_type in mes_sheets:
        if sheet_name not in wb.sheetnames:
            print(f"  ⚠️ Sheet '{sheet_name}' 不存在，跳过")
            continue
        ws = wb[sheet_name]
        count = 0
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not any(row):
                continue
            vals = [clean_text(v) for v in row]
            if len(vals) < 10:
                continue

            device = vals[19] if len(vals) > 19 else ''
            batch = vals[22] if len(vals) > 22 else ''
            b = vals[23] if len(vals) > 23 else ''
            test_program = vals[24] if len(vals) > 24 else ''
            device_prog_bin = f"{device}{test_program}{b}" if device and test_program and b else ''
            # 批次数量
            qty = safe_int(vals[9])

            all_inv_rows.append({
                'device': device,
                'marking': batch,
                'qty': qty,
                'bin': b,
                'test_program': test_program,
                'location_code': wh_name,
                'warehouse_type': wh_type,
                'warehouse_name': wh_name,
                'batch': batch,
                'date_code': vals[20] if len(vals) > 20 else '',
                'material_code': vals[1] if len(vals) > 1 else '',
                'status': vals[12] if len(vals) > 12 else '正常',
                'device_prog_bin': device_prog_bin,
                'import_batch': batch_id,
            })
            count += 1
        print(f"  ✅ {sheet_name}: {count} 条")

    # QHBS 保税仓
    if 'QHBS' in wb.sheetnames:
        ws = wb['QHBS']
        count = 0
        for row in ws.iter_rows(min_row=3, values_only=True):
            if not row or not any(row):
                continue
            vals = [clean_text(v) for v in row]
            if len(vals) < 13:
                continue

            batch_str = vals[12] if len(vals) > 12 else ''
            qty = safe_int(vals[13]) if len(vals) > 13 else 0
            # 解析批次
            device = vals[19] if len(vals) > 19 else ''
            b = vals[18] if len(vals) > 18 else ''
            test_program = vals[17] if len(vals) > 17 else ''
            device_prog_bin = f"{device}{test_program}{b}" if device and test_program and b else ''

            all_inv_rows.append({
                'device': device,
                'marking': batch_str,
                'qty': qty,
                'bin': b,
                'test_program': test_program,
                'location_code': 'QHBS',
                'warehouse_type': 'bonded',
                'warehouse_name': 'QHBS',
                'batch': batch_str,
                'material_code': vals[1] if len(vals) > 1 else '',
                'status': '正常',
                'device_prog_bin': device_prog_bin,
                'import_batch': batch_id,
            })
            count += 1
        print(f"  ✅ QHBS: {count} 条")

    # osat库存
    if 'osat库存' in wb.sheetnames:
        ws = wb['osat库存']
        count = 0
        for row in ws.iter_rows(min_row=3, values_only=True):
            if not row or not any(row):
                continue
            vals = [clean_text(v) for v in row]
            if len(vals) < 6:
                continue
            device = vals[0] or ''
            marking = vals[1] or ''
            b = vals[3] or ''
            test_program = vals[4] or ''
            device_prog_bin = f"{device}{test_program}{b}" if device and test_program and b else ''

            all_inv_rows.append({
                'device': device,
                'marking': marking,
                'qty': safe_int(vals[2]),
                'bin': b,
                'test_program': test_program,
                'location_code': vals[5] if len(vals) > 5 else '',
                'warehouse_type': 'osat',
                'warehouse_name': vals[5] if len(vals) > 5 else '',
                'status': '正常',
                'device_prog_bin': device_prog_bin,
                'import_batch': batch_id,
            })
            count += 1
        print(f"  ✅ osat库存: {count} 条")

    # hold库存
    if 'hold' in wb.sheetnames:
        ws = wb['hold']
        count = 0
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not any(row):
                continue
            vals = [clean_text(v) for v in row]
            if len(vals) < 6:
                continue
            device = vals[0] or ''
            marking = vals[1] or ''
            b = vals[3] or ''
            test_program = vals[4] or ''
            device_prog_bin = f"{device}{test_program}{b}" if device and test_program and b else ''

            all_inv_rows.append({
                'device': device,
                'marking': marking,
                'qty': safe_int(vals[2]),
                'bin': b,
                'test_program': test_program,
                'location_code': vals[5] if len(vals) > 5 else '',
                'warehouse_type': 'hold',
                'warehouse_name': vals[5] if len(vals) > 5 else '',
                'status': 'hold',
                'device_prog_bin': device_prog_bin,
                'import_batch': batch_id,
            })
            count += 1
        print(f"  ✅ hold库存: {count} 条")

    wb.close()

    delete_where('inventory')
    cnt = insert_many('inventory', all_inv_rows)
    print(f"  ✅ 库存总计: {cnt} 条")

def import_kit_completion():
    """导入各外协齐套达成情况"""
    print("\n📋 导入各外协齐套达成情况...")
    import openpyxl
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True, read_only=True)
    ws = wb['各外协齐套达成情况']
    rows = []

    current_region = ''
    current_location = ''
    current_device = ''
    current_model = ''
    current_project = ''
    current_usage = 0
    current_sub = ''
    current_sub_code = ''
    current_initial_stock = 0

    for i, row in enumerate(ws.iter_rows(min_row=1, values_only=True)):
        if not row:
            continue
        vals = [clean_text(v) for v in row]

        # 检测是否是数据行（有芯片和外协信息）
        if vals[0] and vals[3]:
            current_region = vals[0]
            current_location = vals[1] if len(vals) > 1 else ''
            current_device = vals[2] if len(vals) > 2 else ''
            current_model = vals[3] if len(vals) > 3 else ''
            current_project = vals[4] if len(vals) > 4 else ''
            current_usage = safe_int(vals[5]) if len(vals) > 5 else 0
            current_sub = vals[6] if len(vals) > 6 else ''
            current_sub_code = vals[7] if len(vals) > 7 else ''

            # 解析月份计划
            month_plan = {}
            month_names = ['9月','10月','11月','12月','1月','2月','3月','4月','5月','6月','7月','8月']
            for j, m in enumerate(month_names):
                col_idx = 8 + j
                if col_idx < len(vals):
                    month_plan[f'2026-{m}'] = safe_int(vals[col_idx])

            # 期初库存
            stock_col = 8 + len(month_names)
            if stock_col < len(vals):
                current_initial_stock = safe_int(vals[stock_col])

            # 备注
            remark = ''
            remark_col = stock_col + 3
            if remark_col < len(vals):
                remark = vals[remark_col]

            rows.append({
                'region': current_region,
                'location': current_location,
                'device': current_device,
                'model_name': current_model,
                'project': current_project,
                'usage_per_unit': current_usage,
                'subcontractor': current_sub,
                'sub_code': current_sub_code,
                'month_plan': json.dumps(month_plan, ensure_ascii=False),
                'initial_stock': current_initial_stock,
                'current_stock': current_initial_stock,
                'shortage': '{}',
                'actual_ship': '{}',
                'planned_arrival': '{}',
                'remark': remark,
            })
        if i % 500 == 0:
            print(f"  已处理 {i} 行...")

    wb.close()
    delete_where('kit_completion')
    cnt = insert_many('kit_completion', rows)
    print(f"  ✅ 齐套达成: {cnt} 条")

def import_ems_inventory():
    """导入EMS库存（芯片结存统计）"""
    print("\n📋 导入EMS外协库存...")
    if not os.path.exists(EMS_FILE):
        print(f"  ⚠️ EMS文件不存在: {EMS_FILE}")
        return

    import openpyxl
    wb = openpyxl.load_workbook(EMS_FILE, data_only=True, read_only=True)

    if '各外EMS外协库存明细' in wb.sheetnames:
        ws = wb['各外EMS外协库存明细']
        batch_id = generate_batch_id()
        rows = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not any(row):
                continue
            vals = [clean_text(v) for v in row]
            if len(vals) < 8:
                continue
            device = vals[3] if len(vals) > 3 else ''
            b = vals[7] if len(vals) > 7 else ''
            test_program = vals[6] if len(vals) > 6 else ''
            device_prog_bin = f"{device}{test_program}{b}" if device and test_program and b else ''

            rows.append({
                'device': device,
                'marking': vals[5] if len(vals) > 5 else '',
                'qty': safe_int(vals[4]),
                'bin': b,
                'test_program': test_program,
                'location_code': vals[0] if len(vals) > 0 else '',
                'warehouse_type': 'ems',
                'warehouse_name': vals[0] if len(vals) > 0 else '',
                'batch': vals[5] if len(vals) > 5 else '',
                'material_code': vals[1] if len(vals) > 1 else '',
                'status': '正常',
                'device_prog_bin': device_prog_bin,
                'import_batch': batch_id,
            })

        cnt = insert_many('inventory', rows)
        print(f"  ✅ EMS库存: {cnt} 条")
    wb.close()

def main():
    print("=" * 60)
    print("🦞 芯片齐套管理系统 - 历史数据迁移")
    print("=" * 60)

    init_db()

    import_model_mapping()
    import_shipping_detail()
    import_usage_mapping()
    import_mix_bin()
    import_subcontractor_mapping()
    import_logistics()
    import_material_device()
    import_inventory_sheets()
    import_kit_completion()
    import_ems_inventory()

    print("\n" + "=" * 60)
    print("✅ 数据迁移完成！")
    print(f"   出货明细: {count('shipping_detail')} 条")
    print(f"   库存: {count('inventory')} 条")
    print(f"   机型对照: {count('model_mapping')} 条")
    print(f"   用量对照: {count('usage_mapping')} 条")
    print(f"   混BIN: {count('mix_bin')} 条")
    print(f"   齐套达成: {count('kit_completion')} 条")
    print("=" * 60)

if __name__ == "__main__":
    main()
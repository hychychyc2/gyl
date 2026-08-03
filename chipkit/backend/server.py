"""
芯片齐套管理系统 - 精简HTTP服务器
使用Python内置http.server + JSON API，无需FastAPI依赖
"""
import os
import sys
import json
import time
import io
import re
import urllib.parse
from http.server import HTTPServer, BaseHTTPRequestHandler
from datetime import datetime

sys.path.insert(0, os.path.dirname(__file__))
from database import (
    init_db, get_conn, insert, insert_many, update, delete, delete_where,
    query, count, raw, generate_batch_id, write_lock, close_connections
)

FRONTEND_DIR = os.path.join(os.path.dirname(__file__), "..", "frontend")
EXPORTS_DIR = os.path.join(os.path.dirname(__file__), "..", "exports")
TEMP_DIR = os.path.join("/tmp", "chipkit_attachments")
os.makedirs(EXPORTS_DIR, exist_ok=True)
os.makedirs(TEMP_DIR, exist_ok=True)

def parse_post_body(handler):
    """Parse POST body as JSON"""
    content_length = int(handler.headers.get('Content-Length', 0))
    body = handler.rfile.read(content_length)
    if body:
        return json.loads(body.decode('utf-8'))
    return {}

def parse_multipart(handler):
    """Parse multipart form data"""
    content_type = handler.headers.get('Content-Type', '')
    if 'multipart/form-data' not in content_type:
        return {}, {}
    # Extract boundary
    boundary = content_type.split('boundary=')[1].strip()
    body = handler.rfile.read(int(handler.headers.get('Content-Length', 0)))
    parts = body.split(f'--{boundary}'.encode())
    fields = {}
    files = {}
    for part in parts:
        if b'Content-Disposition' not in part:
            continue
        header_end = part.find(b'\r\n\r\n')
        if header_end == -1:
            continue
        headers_raw = part[:header_end].decode('utf-8', errors='ignore')
        content = part[header_end + 4:]
        if content.endswith(b'\r\n'):
            content = content[:-2]

        # Parse name
        name_match = re.search(r'name="([^"]+)"', headers_raw)
        if not name_match:
            continue
        name = name_match.group(1)

        # Check if file
        if 'filename=' in headers_raw:
            fn_match = re.search(r'filename="([^"]*)"', headers_raw)
            filename = fn_match.group(1) if fn_match else 'unknown'
            files[name] = (filename, content)
        else:
            fields[name] = content.decode('utf-8', errors='ignore')

    return fields, files

def send_json(handler, data, status=200):
    handler.send_response(status)
    handler.send_header('Content-Type', 'application/json; charset=utf-8')
    handler.send_header('Access-Control-Allow-Origin', '*')
    handler.end_headers()
    handler.wfile.write(json.dumps(data, ensure_ascii=False).encode('utf-8'))

def send_file(handler, path, content_type=None):
    if not os.path.exists(path):
        handler.send_error(404)
        return
    handler.send_response(200)
    if content_type:
        handler.send_header('Content-Type', content_type)
    handler.send_header('Access-Control-Allow-Origin', '*')
    handler.end_headers()
    with open(path, 'rb') as f:
        handler.wfile.write(f.read())

class ChipKitHandler(BaseHTTPRequestHandler):
    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET,POST,PUT,DELETE,OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def do_GET(self):
        path = urllib.parse.urlparse(self.path).path
        qs = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)

        # API routes
        if path == '/api/dashboard':
            return self.handle_dashboard()
        elif path.startswith('/api/query/'):
            return self.handle_generic_query(path, qs)
        elif path.startswith('/api/count/'):
            return self.handle_count(path, qs)
        elif path.startswith('/api/inventory/'):
            return self.handle_inventory(path, qs)
        elif path.startswith('/api/shipping/'):
            return self.handle_shipping(path, qs)
        elif path.startswith('/api/model/'):
            return self.handle_model(path, qs)
        elif path.startswith('/api/mixbin/'):
            return self.handle_mixbin(path)
        elif path.startswith('/api/kit/'):
            return self.handle_kit(path, qs)
        elif path.startswith('/api/mapping/'):
            return self.handle_mapping(path, qs)
        elif path.startswith('/api/email_configs'):
            return self.handle_email_configs(path, qs)
        elif path.startswith('/api/usage/'):
            return self.handle_usage(path)
        elif path.startswith('/api/users'):
            return self.handle_users()
        elif path.startswith('/api/logs'):
            return self.handle_logs(qs)
        elif path.startswith('/api/export/'):
            return self.handle_export(path)
        elif path.startswith('/api/raw'):
            return self.handle_raw(qs)

        # Static files
        if path == '/' or path == '/index.html':
            return send_file(self, os.path.join(FRONTEND_DIR, 'index.html'), 'text/html')
        filepath = os.path.join(FRONTEND_DIR, path.lstrip('/'))
        if os.path.exists(filepath) and os.path.isfile(filepath):
            ct = 'text/html' if path.endswith('.html') else 'text/javascript' if path.endswith('.js') else 'text/css' if path.endswith('.css') else None
            return send_file(self, filepath, ct)

        send_json(self, {'ok': False, 'error': 'Not found'}, 404)

    def do_POST(self):
        path = urllib.parse.urlparse(self.path).path
        content_type = self.headers.get('Content-Type', '')

        if 'multipart/form-data' in content_type:
            fields, files = parse_multipart(self)
            return self.handle_upload(path, fields, files)

        body = parse_post_body(self)

        if path.startswith('/api/insert/'):
            return self.handle_insert(path, body)
        elif path.startswith('/api/insert_many/'):
            return self.handle_insert_many(path, body)
        elif path.startswith('/api/query/'):
            return self.handle_query_post(path, body)
        elif path.startswith('/api/kit/calculate_shortage'):
            return self.handle_calculate_shortage(body)
        elif path.startswith('/api/shipping/auto_plan'):
            return self.handle_auto_plan()
        elif path.startswith('/api/email_configs/'):
            return self.handle_email_config_post(path, body)
        elif path.startswith('/api/mapping/'):
            return self.handle_mapping_post(path, body)
        elif path.startswith('/api/usage'):
            return self.handle_usage_post(path, body)
        elif path.startswith('/api/users'):
            return self.handle_users_post(path, body)

        send_json(self, {'ok': False, 'error': 'Unknown route'}, 404)

    def do_PUT(self):
        path = urllib.parse.urlparse(self.path).path
        body = parse_post_body(self)

        if path.startswith('/api/update/'):
            return self.handle_update(path, body)
        elif path.startswith('/api/mapping/'):
            return self.handle_mapping_put(path, body)
        elif path.startswith('/api/usage/'):
            return self.handle_usage_put(path, body)
        elif path.startswith('/api/email_configs/'):
            return self.handle_update(path, body)

        send_json(self, {'ok': False, 'error': 'Unknown route'}, 404)

    def do_DELETE(self):
        path = urllib.parse.urlparse(self.path).path

        if path.startswith('/api/delete/'):
            return self.handle_delete(path)
        elif path.startswith('/api/mapping/'):
            return self.handle_delete(path)
        elif path.startswith('/api/email_configs/'):
            return self.handle_delete(path)

        send_json(self, {'ok': False, 'error': 'Unknown route'}, 404)

    # ============ Handlers ============
    def handle_dashboard(self):
        data = {
            'total_shipping': count('shipping_detail'),
            'total_inventory': count('inventory'),
            'total_models': count('model_mapping'),
            'total_kit': count('kit_completion'),
            'inventory_by_type': raw("SELECT warehouse_type, SUM(qty) as total_qty FROM inventory GROUP BY warehouse_type"),
            'recent_shipping': query('shipping_detail', order_by='ship_date DESC', limit=10),
            'low_stock': raw("SELECT device, SUM(qty) as total_qty FROM inventory WHERE warehouse_type='osat' GROUP BY device HAVING total_qty < 1000 ORDER BY total_qty"),
        }
        send_json(self, {'ok': True, 'data': data})

    def handle_generic_query(self, path, qs):
        parts = path.split('/')
        table = parts[3]
        where = qs.get('where', [''])[0]
        params = json.loads(qs.get('params', ['[]'])[0])
        order_by = qs.get('order_by', [''])[0]
        limit = int(qs.get('limit', ['0'])[0])
        offset = int(qs.get('offset', ['0'])[0])
        try:
            rows = query(table, where=where, params=tuple(params), order_by=order_by, limit=limit, offset=offset)
            total = count(table, where=where, params=tuple(params))
            send_json(self, {'ok': True, 'data': rows, 'total': total})
        except Exception as e:
            send_json(self, {'ok': False, 'error': str(e)})

    def handle_query_post(self, path, body):
        parts = path.split('/')
        table = parts[3]
        try:
            rows = query(table, where=body.get('where', ''), params=tuple(body.get('params', [])),
                         order_by=body.get('order_by', ''), limit=body.get('limit', 0),
                         offset=body.get('offset', 0), columns=body.get('columns', '*'))
            total = count(table, where=body.get('where', ''), params=tuple(body.get('params', [])))
            send_json(self, {'ok': True, 'data': rows, 'total': total})
        except Exception as e:
            send_json(self, {'ok': False, 'error': str(e)})

    def handle_count(self, path, qs):
        parts = path.split('/')
        table = parts[3]
        where = qs.get('where', [''])[0]
        params = json.loads(qs.get('params', ['[]'])[0])
        try:
            c = count(table, where=where, params=tuple(params))
            send_json(self, {'ok': True, 'count': c})
        except Exception as e:
            send_json(self, {'ok': False, 'error': str(e)})

    def handle_insert(self, path, body):
        parts = path.split('/')
        table = parts[3]
        try:
            rid = insert(table, body)
            send_json(self, {'ok': True, 'id': rid})
        except Exception as e:
            send_json(self, {'ok': False, 'error': str(e)})

    def handle_insert_many(self, path, body):
        parts = path.split('/')
        table = parts[3]
        try:
            cnt = insert_many(table, body)
            send_json(self, {'ok': True, 'count': cnt})
        except Exception as e:
            send_json(self, {'ok': False, 'error': str(e)})

    def handle_update(self, path, body):
        parts = path.split('/')
        # /api/update/{table}/{id}
        if len(parts) >= 5:
            table = parts[3]
            record_id = int(parts[4])
            try:
                ok = update(table, record_id, body)
                send_json(self, {'ok': ok})
            except Exception as e:
                send_json(self, {'ok': False, 'error': str(e)})
        else:
            send_json(self, {'ok': False, 'error': 'Invalid path'}, 400)

    def handle_delete(self, path):
        parts = path.split('/')
        # /api/delete/{table}/{id} or /api/mapping/{table}/{id} or /api/email_configs/{id}
        try:
            if parts[1] == 'delete':
                table = parts[2]
                record_id = int(parts[3])
            elif parts[1] == 'mapping':
                table = parts[2]
                record_id = int(parts[3])
            elif parts[1] == 'email_configs':
                table = 'email_config'
                record_id = int(parts[2])
            else:
                send_json(self, {'ok': False, 'error': 'Invalid path'}, 400)
                return
            ok = delete(table, record_id)
            send_json(self, {'ok': ok})
        except Exception as e:
            send_json(self, {'ok': False, 'error': str(e)})

    def handle_inventory(self, path, qs):
        if path == '/api/inventory/summary':
            rows = raw("SELECT warehouse_type, warehouse_name, COUNT(*) as batch_count, SUM(qty) as total_qty FROM inventory GROUP BY warehouse_type, warehouse_name ORDER BY warehouse_type, total_qty DESC")
            send_json(self, {'ok': True, 'data': rows})
        elif path == '/api/inventory/with_model':
            device = qs.get('device', [''])[0]
            where = "WHERE i.device = ?" if device else ""
            params = [device] if device else []
            sql = f"""
            SELECT i.device, i.device_prog_bin, i.bin, i.test_program, i.warehouse_name, i.warehouse_type,
                   SUM(i.qty) as total_qty, m.model1, m.model2, m.model3, u.usage_qty
            FROM inventory i
            LEFT JOIN model_mapping m ON i.device_prog_bin = m.device_prog_bin
            LEFT JOIN usage_mapping u ON i.device = u.device AND (m.model1 = u.model_name OR m.model2 = u.model_name)
            {where}
            GROUP BY i.device_prog_bin, i.warehouse_name, i.warehouse_type
            ORDER BY i.device, total_qty DESC
            """
            rows = raw(sql, tuple(params))
            for r in rows:
                qty = r.get('total_qty', 0) or 0
                usage = r.get('usage_qty', 0) or 0
                r['machine_count'] = int(qty // usage) if usage > 0 else 0
            send_json(self, {'ok': True, 'data': rows})
        else:
            send_json(self, {'ok': False, 'error': 'Unknown route'}, 404)

    def handle_shipping(self, path, qs):
        if path == '/api/shipping/expired':
            rows = raw("SELECT * FROM shipping_detail WHERE shipped_qty = 0 AND ship_date < date('now', '-180 days') ORDER BY ship_date DESC")
            send_json(self, {'ok': True, 'data': rows, 'total': len(rows)})
        else:
            send_json(self, {'ok': False, 'error': 'Unknown route'}, 404)

    def handle_model(self, path, qs):
        if path == '/api/model/mapping':
            device = qs.get('device', [''])[0]
            model_name = qs.get('model_name', [''])[0]
            where = []
            params = []
            if device:
                where.append('device LIKE ?')
                params.append(f'%{device}%')
            if model_name:
                where.append('(model1 LIKE ? OR model2 LIKE ?)')
                params.extend([f'%{model_name}%', f'%{model_name}%'])
            sql = "SELECT * FROM model_mapping"
            if where:
                sql += " WHERE " + " AND ".join(where)
            sql += " ORDER BY device, bin LIMIT 500"
            rows = raw(sql, tuple(params))
            send_json(self, {'ok': True, 'data': rows, 'total': len(rows)})
        elif path == '/api/model/with_stock':
            rows = raw("""
            SELECT m.*, COALESCE(i.total_qty, 0) as stock_qty, u.usage_qty,
                   CASE WHEN u.usage_qty > 0 THEN COALESCE(i.total_qty, 0) / u.usage_qty ELSE 0 END as machine_count
            FROM model_mapping m
            LEFT JOIN (SELECT device_prog_bin, SUM(qty) as total_qty FROM inventory GROUP BY device_prog_bin) i ON m.device_prog_bin = i.device_prog_bin
            LEFT JOIN usage_mapping u ON m.device = u.device AND m.model1 = u.model_name
            ORDER BY m.device, m.bin LIMIT 500
            """)
            send_json(self, {'ok': True, 'data': rows})
        else:
            send_json(self, {'ok': False, 'error': 'Unknown route'}, 404)

    def handle_mixbin(self, path):
        rows = query('mix_bin', order_by='device, bin')
        send_json(self, {'ok': True, 'data': rows})

    def handle_kit(self, path, qs):
        if path == '/api/kit/completion':
            region = qs.get('region', [''])[0]
            where = "region = ?" if region else ""
            params = (region,) if region else ()
            rows = query('kit_completion', where=where, params=params, order_by='region, location, device, subcontractor')
            for r in rows:
                for field in ['month_plan', 'shortage', 'actual_ship', 'planned_arrival']:
                    try:
                        r[field] = json.loads(r.get(field, '{}') or '{}')
                    except:
                        r[field] = {}
            send_json(self, {'ok': True, 'data': rows})
        else:
            send_json(self, {'ok': False, 'error': 'Unknown route'}, 404)

    def handle_calculate_shortage(self, body):
        region = body.get('region', '')
        where = 'region = ?' if region else ''
        params = (region,) if region else ()
        rows = query('kit_completion', where=where, params=params)
        results = []
        for r in rows:
            month_plan = json.loads(r.get('month_plan', '{}') or '{}')
            usage = r.get('usage_per_unit', 0) or 0
            initial_stock = r.get('initial_stock', 0) or 0
            shortage = {}
            cumulative = initial_stock
            for month, plan in sorted(month_plan.items()):
                needed = int(plan) * usage if usage else 0
                cumulative -= needed
                shortage[month] = cumulative
            update('kit_completion', r['id'], {
                'shortage': json.dumps(shortage, ensure_ascii=False),
                'version': r.get('version', 1)
            })
            results.append({**r, 'shortage': shortage})
        send_json(self, {'ok': True, 'data': results})

    def handle_auto_plan(self):
        kits = query('kit_completion')
        plans = []
        for kit in kits:
            shortage = json.loads(kit.get('shortage', '{}') or '{}')
            total = sum(abs(v) for v in shortage.values() if v < 0)
            if total <= 0:
                continue
            device = kit.get('device', '')
            stocks = query('inventory', where='device=? AND warehouse_type IN (?,?,?)',
                          params=(device, 'osat', 'bonded', 'other'),
                          order_by='CASE warehouse_type WHEN "osat" THEN 1 WHEN "bonded" THEN 2 ELSE 3 END')
            remaining = total
            for stock in stocks:
                if remaining <= 0: break
                qty = min(stock.get('qty', 0), remaining)
                if qty <= 0: continue
                plans.append({
                    'plan_date': datetime.now().strftime('%Y-%m-%d'),
                    'osat': stock.get('warehouse_name', ''),
                    'device': device, 'bin': stock.get('bin', ''), 'qty': qty,
                    'from_warehouse': stock.get('warehouse_name', ''),
                    'warehouse_type': stock.get('warehouse_type', ''),
                    'ship_to': kit.get('subcontractor', ''),
                    'model_name': kit.get('model_name', ''),
                    'project': kit.get('project', ''), 'status': '待确认',
                })
                remaining -= qty
        if plans:
            delete_where('shipping_plan')
            insert_many('shipping_plan', plans)
        send_json(self, {'ok': True, 'data': plans, 'total': len(plans)})

    def handle_mapping(self, path, qs):
        parts = path.split('/')
        table = parts[2] if len(parts) > 2 else ''
        valid = {'subcontractor_mapping', 'logistics_time', 'material_device'}
        if table not in valid:
            send_json(self, {'ok': False, 'error': 'Invalid table'}, 400)
            return
        search = qs.get('search', [''])[0]
        where = ''
        params = []
        if search and table == 'subcontractor_mapping':
            where = 'short_name LIKE ? OR internal_code LIKE ? OR external_name LIKE ?'
            params = [f'%{search}%'] * 3
        elif search and table == 'material_device':
            where = 'erp_code LIKE ? OR device LIKE ?'
            params = [f'%{search}%'] * 2
        elif search and table == 'logistics_time':
            where = 'destination LIKE ?'
            params = [f'%{search}%']
        rows = query(table, where=where, params=tuple(params))
        send_json(self, {'ok': True, 'data': rows})

    def handle_mapping_post(self, path, body):
        parts = path.split('/')
        table = parts[2]
        valid = {'subcontractor_mapping', 'logistics_time', 'material_device'}
        if table not in valid:
            send_json(self, {'ok': False, 'error': 'Invalid table'}, 400)
            return
        rid = insert(table, body)
        send_json(self, {'ok': True, 'id': rid})

    def handle_mapping_put(self, path, body):
        parts = path.split('/')
        table = parts[2]
        record_id = int(parts[3])
        ok = update(table, record_id, body)
        send_json(self, {'ok': ok})

    def handle_email_configs(self, path, qs):
        rows = query('email_config', order_by='purpose')
        for r in rows:
            if r.get('password_blob'):
                r['password_blob'] = '********'
        send_json(self, {'ok': True, 'data': rows})

    def handle_email_config_post(self, path, body):
        if len(path.split('/')) > 3:
            # /api/email_configs/{id}/fetch
            record_id = int(path.split('/')[2])
            try:
                from email_collector import download_email_attachments, process_shipping_detail, process_inventory, process_model_mapping, process_mix_bin, process_order_allocation
                cfg = query('email_config', where='id=?', params=(record_id,))
                if not cfg:
                    return send_json(self, {'ok': False, 'error': '配置不存在'})
                cfg = dict(cfg[0])
                file_path = download_email_attachments(cfg, TEMP_DIR)
                if not file_path:
                    return send_json(self, {'ok': False, 'error': '未找到附件'})
                count = 0
                purpose = cfg.get('purpose', '')
                if purpose == 'shipping_detail':
                    count = process_shipping_detail(file_path, cfg)
                elif purpose == 'inventory':
                    count = process_inventory(file_path, cfg)
                elif purpose == 'model_mapping':
                    count = process_model_mapping(file_path, cfg)
                elif purpose == 'mix_bin':
                    count = process_mix_bin(file_path, cfg)
                elif purpose == 'order_allocation':
                    count = process_order_allocation(file_path, cfg)
                update('email_config', record_id, {
                    'last_fetch': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'version': cfg.get('version', 1)
                })
                return send_json(self, {'ok': True, 'count': count})
            except Exception as e:
                return send_json(self, {'ok': False, 'error': str(e)})
        else:
            rid = insert('email_config', body)
            return send_json(self, {'ok': True, 'id': rid})

    def handle_upload(self, path, fields, files):
        if path == '/api/upload/inventory':
            if 'file' not in files:
                return send_json(self, {'ok': False, 'error': 'No file'})
            filename, content = files['file']
            filepath = os.path.join(TEMP_DIR, f"upload_{int(time.time())}_{filename}")
            with open(filepath, 'wb') as f:
                f.write(content)
            try:
                from email_collector import parse_excel, generate_batch_id
                warehouse_type = fields.get('warehouse_type', 'other')
                warehouse_name = fields.get('warehouse_name', '')
                sheet_name = fields.get('sheet_name', 'Sheet1')
                header_row = int(fields.get('header_row', '1'))
                rows = parse_excel(filepath, sheet_name, header_row)
                if not rows:
                    return send_json(self, {'ok': False, 'error': '文件解析无数据'})
                batch_id = generate_batch_id()
                inv_rows = []
                for row in rows:
                    device = row.get('device', '')
                    marking = row.get('marking', '')
                    b = row.get('bin', '')
                    tp = row.get('test_program', '')
                    dpb = f"{device}{tp}{b}" if device and tp and b else ''
                    inv_rows.append({
                        'device': device, 'marking': marking,
                        'qty': int(row.get('qty', 0) or 0),
                        'bin': b, 'test_program': tp,
                        'location_code': row.get('location_code', warehouse_name),
                        'warehouse_type': warehouse_type, 'warehouse_name': warehouse_name,
                        'batch': row.get('batch', ''), 'date_code': row.get('date_code', ''),
                        'material_code': row.get('material_code', ''),
                        'status': row.get('status', '正常'),
                        'device_prog_bin': dpb, 'import_batch': batch_id,
                    })
                cnt = insert_many('inventory', inv_rows)
                return send_json(self, {'ok': True, 'count': cnt, 'batch_id': batch_id})
            finally:
                try: os.remove(filepath)
                except: pass
        elif path == '/api/upload/shipping':
            if 'file' not in files:
                return send_json(self, {'ok': False, 'error': 'No file'})
            filename, content = files['file']
            filepath = os.path.join(TEMP_DIR, f"ship_{int(time.time())}_{filename}")
            with open(filepath, 'wb') as f:
                f.write(content)
            try:
                from email_collector import parse_excel, generate_batch_id
                rows = parse_excel(filepath, 'Sheet1', 1)
                batch_id = generate_batch_id()
                ship_rows = []
                for row in rows:
                    ship_rows.append({
                        'entity': row.get('entity', ''), 'ship_date': row.get('ship_date', ''),
                        'device_pn': row.get('device_pn', ''), 'wafer_lot_id': row.get('wafer_lot_id', ''),
                        'marking': row.get('marking', ''),
                        'good_qty': int(row.get('good_qty', 0) or 0),
                        'bin': row.get('bin', ''), 'invoice_no': row.get('invoice_no', ''),
                        'test_program': row.get('test_program', ''), 'osat': row.get('osat', ''),
                        'ship_to': row.get('ship_to', ''), 'test_wo': row.get('test_wo', ''),
                        'date_code': row.get('date_code', ''),
                        'po': row.get('po', '暂无'), 'source': 'upload', 'import_batch': batch_id,
                    })
                cnt = insert_many('shipping_detail', ship_rows)
                return send_json(self, {'ok': True, 'count': cnt, 'batch_id': batch_id})
            finally:
                try: os.remove(filepath)
                except: pass
        elif path == '/api/upload/erp_inventory':
            if 'file' not in files:
                return send_json(self, {'ok': False, 'error': 'No file'})
            filename, content = files['file']
            filepath = os.path.join(TEMP_DIR, f"erp_{int(time.time())}_{filename}")
            with open(filepath, 'wb') as f:
                f.write(content)
            try:
                from email_collector import parse_excel, generate_batch_id
                rows = parse_excel(filepath, 'Sheet1', 1)
                batch_id = generate_batch_id()
                erp_rows = []
                for row in rows:
                    device = row.get('device', '')
                    b = row.get('bin', '')
                    tp = row.get('test_program', '')
                    dpb = f"{device}{tp}{b}" if device and tp and b else ''
                    erp_rows.append({
                        'org': row.get('org', ''), 'material_code': row.get('material_code', ''),
                        'description': row.get('description', ''),
                        'sub_inventory': row.get('sub_inventory', ''),
                        'location': row.get('location', ''),
                        'batch': row.get('batch', ''),
                        'qty': int(row.get('qty', 0) or 0),
                        'device': device, 'bin': b, 'test_program': tp,
                        'device_prog_bin': dpb, 'import_batch': batch_id,
                    })
                cnt = insert_many('erp_inventory', erp_rows)
                return send_json(self, {'ok': True, 'count': cnt, 'batch_id': batch_id})
            finally:
                try: os.remove(filepath)
                except: pass
        send_json(self, {'ok': False, 'error': 'Unknown upload route'}, 404)

    def handle_usage(self, path):
        rows = query('usage_mapping', order_by='device, model_name')
        send_json(self, {'ok': True, 'data': rows})

    def handle_usage_post(self, path, body):
        rid = insert('usage_mapping', body)
        send_json(self, {'ok': True, 'id': rid})

    def handle_usage_put(self, path, body):
        parts = path.split('/')
        record_id = int(parts[2])
        ok = update('usage_mapping', record_id, body)
        send_json(self, {'ok': ok})

    def handle_users(self):
        rows = query('users', columns='id, email, name, role, active, created_at')
        send_json(self, {'ok': True, 'data': rows})

    def handle_users_post(self, path, body):
        rid = insert('users', body)
        send_json(self, {'ok': True, 'id': rid})

    def handle_logs(self, qs):
        limit = int(qs.get('limit', ['100'])[0])
        rows = query('operation_log', order_by='created_at DESC', limit=limit)
        send_json(self, {'ok': True, 'data': rows})

    def handle_export(self, path):
        parts = path.split('/')
        table = parts[3]
        try:
            rows = query(table, limit=50000)
            filepath = os.path.join(EXPORTS_DIR, f"{table}_{datetime.now().strftime('%Y%m%d')}.json")
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(rows, f, ensure_ascii=False, indent=2)
            send_file(self, filepath, 'application/json')
        except Exception as e:
            send_json(self, {'ok': False, 'error': str(e)})

    def handle_raw(self, qs):
        sql = qs.get('sql', [''])[0]
        params = json.loads(qs.get('params', ['[]'])[0])
        fetch = qs.get('fetch', ['true'])[0] == 'true'
        try:
            if fetch:
                rows = raw(sql, tuple(params), fetch=True)
                send_json(self, {'ok': True, 'data': rows})
            else:
                cnt = raw(sql, tuple(params), fetch=False)
                send_json(self, {'ok': True, 'count': cnt})
        except Exception as e:
            send_json(self, {'ok': False, 'error': str(e)})

    def log_message(self, format, *args):
        pass  # Quiet

def main():
    init_db()
    port = 8765
    server = HTTPServer(('0.0.0.0', port), ChipKitHandler)
    print(f"🦞 芯片齐套管理系统启动: http://localhost:{port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        server.shutdown()

if __name__ == '__main__':
    main()
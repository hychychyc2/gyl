"""
芯片齐套管理系统 - 数据库层
SQLite WAL模式 + 应用层乐观锁 + 密码加密
"""
import sqlite3
import threading
import time
import json
import os
import base64
import hashlib
from datetime import datetime
from contextlib import contextmanager
from typing import Optional, List, Dict, Any, Tuple

DB_PATH = os.path.join(os.path.dirname(__file__), "..", "data", "chipkit.db")

_write_lock = threading.Lock()
_local = threading.local()

# ============ 密码加密 ============
SECRET_KEY = b'chipkit_salt_2026'

def encrypt_password(plain: str) -> str:
    """简单加密存储邮箱密码"""
    if not plain:
        return ''
    # 使用固定密钥 XOR + base64
    key = SECRET_KEY * (len(plain) // len(SECRET_KEY) + 1)
    encrypted = bytes(a ^ b for a, b in zip(plain.encode(), key[:len(plain)]))
    return base64.b64encode(encrypted).decode()

def decrypt_password(encrypted: str) -> str:
    """解密邮箱密码"""
    if not encrypted:
        return ''
    try:
        data = base64.b64decode(encrypted)
        key = SECRET_KEY * (len(data) // len(SECRET_KEY) + 1)
        return bytes(a ^ b for a, b in zip(data, key[:len(data)])).decode()
    except:
        return encrypted

# ============ 连接管理 ============
def get_db_path():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    return DB_PATH

def get_conn() -> sqlite3.Connection:
    if not hasattr(_local, 'conn') or _local.conn is None:
        path = get_db_path()
        _local.conn = sqlite3.connect(path, check_same_thread=False)
        _local.conn.row_factory = sqlite3.Row
        _local.conn.execute("PRAGMA journal_mode=WAL")
        _local.conn.execute("PRAGMA synchronous=NORMAL")
        _local.conn.execute("PRAGMA busy_timeout=5000")
        _local.conn.execute("PRAGMA cache_size=-20000")
        _local.conn.execute("PRAGMA foreign_keys=ON")
    return _local.conn

@contextmanager
def write_lock():
    acquired = _write_lock.acquire(timeout=5)
    if not acquired:
        raise RuntimeError("写入锁获取超时")
    try:
        yield
    finally:
        _write_lock.release()

# ============ 表结构 ============
def init_db():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = get_conn()
    cursor = conn.cursor()

    tables = [
        # ===== 出货明细 =====
        """CREATE TABLE IF NOT EXISTS shipping_detail (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            entity TEXT DEFAULT '',
            ship_date TEXT DEFAULT '',
            device_pn TEXT DEFAULT '',
            wafer_lot_id TEXT DEFAULT '',
            marking TEXT DEFAULT '',
            good_qty INTEGER DEFAULT 0,
            bin TEXT DEFAULT '',
            invoice_no TEXT DEFAULT '',
            test_program TEXT DEFAULT '',
            osat TEXT DEFAULT '',
            ship_to TEXT DEFAULT '',
            test_wo TEXT DEFAULT '',
            date_code TEXT DEFAULT '',
            po TEXT DEFAULT '',
            so TEXT DEFAULT '',
            model1 TEXT DEFAULT '',
            line_no TEXT DEFAULT '',
            material_code TEXT DEFAULT '',
            production_code TEXT DEFAULT '',
            order_qty INTEGER DEFAULT 0,
            sub_inventory TEXT DEFAULT '',
            location TEXT DEFAULT '',
            device_prog_bin TEXT DEFAULT '',
            erp_batch_ref TEXT DEFAULT '',
            shipped_qty INTEGER DEFAULT 0,
            source TEXT DEFAULT '',
            import_batch TEXT DEFAULT '',
            version INTEGER DEFAULT 1,
            created_at TEXT DEFAULT (datetime('now','localtime')),
            updated_at TEXT DEFAULT (datetime('now','localtime'))
        )""",
        """CREATE UNIQUE INDEX IF NOT EXISTS idx_ship_unique 
        ON shipping_detail(entity, ship_date, device_pn, wafer_lot_id, marking, 
                          good_qty, bin, invoice_no, test_program, osat, ship_to, test_wo)""",

        # ===== 库存统一表（所有库存类型共用） =====
        """CREATE TABLE IF NOT EXISTS inventory (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            device TEXT DEFAULT '',
            marking TEXT DEFAULT '',
            qty INTEGER DEFAULT 0,
            bin TEXT DEFAULT '',
            test_program TEXT DEFAULT '',
            warehouse_type TEXT DEFAULT '',
            warehouse_name TEXT DEFAULT '',
            material_code TEXT DEFAULT '',
            product_desc TEXT DEFAULT '',
            batch TEXT DEFAULT '',
            date_code TEXT DEFAULT '',
            status TEXT DEFAULT '正常',
            location_area TEXT DEFAULT '',
            sub_inventory TEXT DEFAULT '',
            org TEXT DEFAULT '',
            remark TEXT DEFAULT '',
            device_prog_bin TEXT DEFAULT '',
            import_batch TEXT DEFAULT '',
            version INTEGER DEFAULT 1,
            created_at TEXT DEFAULT (datetime('now','localtime')),
            updated_at TEXT DEFAULT (datetime('now','localtime'))
        )""",
        """CREATE INDEX IF NOT EXISTS idx_inv_device ON inventory(device)""",
        """CREATE INDEX IF NOT EXISTS idx_inv_type ON inventory(warehouse_type)""",
        """CREATE INDEX IF NOT EXISTS idx_inv_dpb ON inventory(device_prog_bin)""",

        # ===== 机型对照表 =====
        """CREATE TABLE IF NOT EXISTS model_mapping (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            material_code TEXT DEFAULT '',
            device TEXT DEFAULT '',
            test_program TEXT DEFAULT '',
            bin TEXT DEFAULT '',
            model1 TEXT DEFAULT '', model2 TEXT DEFAULT '', model3 TEXT DEFAULT '',
            model4 TEXT DEFAULT '', model5 TEXT DEFAULT '', model6 TEXT DEFAULT '',
            model7 TEXT DEFAULT '', model8 TEXT DEFAULT '', model9 TEXT DEFAULT '',
            model10 TEXT DEFAULT '', model11 TEXT DEFAULT '', model12 TEXT DEFAULT '',
            model13 TEXT DEFAULT '', model14 TEXT DEFAULT '', model15 TEXT DEFAULT '',
            model16 TEXT DEFAULT '', model17 TEXT DEFAULT '', model18 TEXT DEFAULT '',
            model19 TEXT DEFAULT '', model20 TEXT DEFAULT '', model21 TEXT DEFAULT '',
            model22 TEXT DEFAULT '',
            device_prog_bin TEXT DEFAULT '',
            product TEXT DEFAULT '',
            osat_model TEXT DEFAULT '',
            project TEXT DEFAULT '',
            exclusive_bin INTEGER DEFAULT 0,
            version INTEGER DEFAULT 1,
            created_at TEXT DEFAULT (datetime('now','localtime')),
            updated_at TEXT DEFAULT (datetime('now','localtime'))
        )""",
        """CREATE UNIQUE INDEX IF NOT EXISTS idx_model_unique 
        ON model_mapping(device, test_program, bin)""",

        # ===== 用量对照表 =====
        """CREATE TABLE IF NOT EXISTS usage_mapping (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            device TEXT DEFAULT '',
            model_name TEXT DEFAULT '',
            project TEXT DEFAULT '',
            usage_qty INTEGER DEFAULT 0,
            version INTEGER DEFAULT 1,
            created_at TEXT DEFAULT (datetime('now','localtime')),
            updated_at TEXT DEFAULT (datetime('now','localtime'))
        )""",

        # ===== 混BIN分配表 =====
        """CREATE TABLE IF NOT EXISTS mix_bin (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            device_prog_bin TEXT DEFAULT '',
            material_code TEXT DEFAULT '',
            device TEXT DEFAULT '',
            test_program TEXT DEFAULT '',
            bin TEXT DEFAULT '',
            col TEXT DEFAULT '',
            model_name TEXT DEFAULT '',
            mix_group TEXT DEFAULT '',
            stock_qty INTEGER DEFAULT 0,
            chips_per_unit INTEGER DEFAULT 0,
            convertible_qty REAL DEFAULT 0,
            summary_actual INTEGER DEFAULT 0,
            is_exclusive INTEGER DEFAULT 0,
            version INTEGER DEFAULT 1,
            created_at TEXT DEFAULT (datetime('now','localtime')),
            updated_at TEXT DEFAULT (datetime('now','localtime'))
        )""",

        # ===== 外协代码映射表 =====
        """CREATE TABLE IF NOT EXISTS subcontractor_mapping (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            type TEXT DEFAULT '',
            short_name TEXT DEFAULT '',
            internal_code TEXT DEFAULT '',
            external_name TEXT DEFAULT '',
            ship_to_code TEXT DEFAULT '',
            address TEXT DEFAULT '',
            contact TEXT DEFAULT '',
            version INTEGER DEFAULT 1,
            created_at TEXT DEFAULT (datetime('now','localtime')),
            updated_at TEXT DEFAULT (datetime('now','localtime'))
        )""",

        # ===== 物流时间表 =====
        """CREATE TABLE IF NOT EXISTS logistics_time (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            destination TEXT DEFAULT '',
            transit_days INTEGER DEFAULT 0,
            latest_ship_day TEXT DEFAULT '',
            version INTEGER DEFAULT 1,
            created_at TEXT DEFAULT (datetime('now','localtime')),
            updated_at TEXT DEFAULT (datetime('now','localtime'))
        )""",

        # ===== 料号Device对照表 =====
        """CREATE TABLE IF NOT EXISTS material_device (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            erp_code TEXT DEFAULT '',
            device TEXT DEFAULT '',
            wafer_pn TEXT DEFAULT '',
            description TEXT DEFAULT '',
            package_desc TEXT DEFAULT '',
            version INTEGER DEFAULT 1,
            created_at TEXT DEFAULT (datetime('now','localtime')),
            updated_at TEXT DEFAULT (datetime('now','localtime'))
        )""",

        # ===== 齐套达成表 =====
        """CREATE TABLE IF NOT EXISTS kit_completion (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            region TEXT DEFAULT '',
            location TEXT DEFAULT '',
            device TEXT DEFAULT '',
            model_name TEXT DEFAULT '',
            project TEXT DEFAULT '',
            usage_per_unit INTEGER DEFAULT 0,
            subcontractor TEXT DEFAULT '',
            sub_code TEXT DEFAULT '',
            month_plan TEXT DEFAULT '{}',
            initial_stock INTEGER DEFAULT 0,
            shortage TEXT DEFAULT '{}',
            actual_ship TEXT DEFAULT '{}',
            planned_arrival TEXT DEFAULT '{}',
            current_stock INTEGER DEFAULT 0,
            remark TEXT DEFAULT '',
            version INTEGER DEFAULT 1,
            created_at TEXT DEFAULT (datetime('now','localtime')),
            updated_at TEXT DEFAULT (datetime('now','localtime'))
        )""",

        # ===== 出货计划表 =====
        """CREATE TABLE IF NOT EXISTS shipping_plan (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            plan_date TEXT DEFAULT '',
            osat TEXT DEFAULT '',
            device TEXT DEFAULT '',
            bin TEXT DEFAULT '',
            qty INTEGER DEFAULT 0,
            from_warehouse TEXT DEFAULT '',
            warehouse_type TEXT DEFAULT '',
            ship_to TEXT DEFAULT '',
            model_name TEXT DEFAULT '',
            project TEXT DEFAULT '',
            remark TEXT DEFAULT '',
            status TEXT DEFAULT '待确认',
            version INTEGER DEFAULT 1,
            created_at TEXT DEFAULT (datetime('now','localtime')),
            updated_at TEXT DEFAULT (datetime('now','localtime'))
        )""",

        # ===== ERP库存表 =====
        """CREATE TABLE IF NOT EXISTS erp_inventory (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            org TEXT DEFAULT '',
            material_code TEXT DEFAULT '',
            description TEXT DEFAULT '',
            sub_inventory TEXT DEFAULT '',
            location TEXT DEFAULT '',
            batch TEXT DEFAULT '',
            qty INTEGER DEFAULT 0,
            device TEXT DEFAULT '',
            bin TEXT DEFAULT '',
            test_program TEXT DEFAULT '',
            device_prog_bin TEXT DEFAULT '',
            import_batch TEXT DEFAULT '',
            version INTEGER DEFAULT 1,
            created_at TEXT DEFAULT (datetime('now','localtime')),
            updated_at TEXT DEFAULT (datetime('now','localtime'))
        )""",

        # ===== 邮件配置表 =====
        """CREATE TABLE IF NOT EXISTS email_config (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            purpose TEXT DEFAULT '',
            description TEXT DEFAULT '',
            email_address TEXT DEFAULT '',
            imap_server TEXT DEFAULT 'imap.appia.vip',
            account TEXT DEFAULT '',
            password_encrypted TEXT DEFAULT '',
            root_folder TEXT DEFAULT 'INBOX',
            match_key TEXT DEFAULT '',
            suffix TEXT DEFAULT '.xlsx',
            mapping_config TEXT DEFAULT '{}',
            active INTEGER DEFAULT 1,
            last_fetch TEXT DEFAULT '',
            version INTEGER DEFAULT 1,
            created_at TEXT DEFAULT (datetime('now','localtime')),
            updated_at TEXT DEFAULT (datetime('now','localtime'))
        )""",

        # ===== 用户表 =====
        """CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT UNIQUE NOT NULL,
            name TEXT DEFAULT '',
            role TEXT DEFAULT 'viewer',
            password_hash TEXT DEFAULT '',
            active INTEGER DEFAULT 1,
            created_at TEXT DEFAULT (datetime('now','localtime'))
        )""",

        # ===== 操作日志 =====
        """CREATE TABLE IF NOT EXISTS operation_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            action TEXT DEFAULT '',
            table_name TEXT DEFAULT '',
            record_id INTEGER DEFAULT 0,
            detail TEXT DEFAULT '',
            operator TEXT DEFAULT '',
            created_at TEXT DEFAULT (datetime('now','localtime'))
        )""",

        # ===== 导入批次 =====
        """CREATE TABLE IF NOT EXISTS import_batch (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            batch_id TEXT UNIQUE DEFAULT '',
            table_name TEXT DEFAULT '',
            source TEXT DEFAULT '',
            file_name TEXT DEFAULT '',
            row_count INTEGER DEFAULT 0,
            status TEXT DEFAULT '完成',
            created_at TEXT DEFAULT (datetime('now','localtime'))
        )""",
    ]

    for sql in tables:
        try:
            cursor.execute(sql)
        except Exception as e:
            print(f"SQL error: {e}")

    conn.commit()

    # 自动迁移：添加版本更新中新增的列
    migrations = {
        'inventory': ['source_email', 'source_file', 'source_time'],
        'shipping_detail': ['source_email', 'source_file', 'source_time'],
    }
    for table, cols in migrations.items():
        existing = {r['name'] for r in conn.execute(f'PRAGMA table_info({table})').fetchall()}
        for c in cols:
            if c not in existing:
                conn.execute(f'ALTER TABLE {table} ADD COLUMN {c} TEXT DEFAULT ""')
    conn.commit()

    print("✅ 数据库初始化完成")

# ============ 通用 CRUD ============
def insert(table: str, data: Dict[str, Any]) -> int:
    with write_lock():
        conn = get_conn()
        columns = list(data.keys())
        placeholders = ', '.join(['?'] * len(columns))
        cols_str = ', '.join(columns)
        sql = f"INSERT INTO {table} ({cols_str}) VALUES ({placeholders})"
        cursor = conn.execute(sql, list(data.values()))
        conn.commit()
        return cursor.lastrowid

def insert_many(table: str, rows: List[Dict[str, Any]]) -> int:
    if not rows:
        return 0
    with write_lock():
        conn = get_conn()
        columns = list(rows[0].keys())
        placeholders = ', '.join(['?'] * len(columns))
        cols_str = ', '.join(columns)
        sql = f"INSERT OR IGNORE INTO {table} ({cols_str}) VALUES ({placeholders})"
        params = [tuple(r.get(c, '') for c in columns) for r in rows]
        cursor = conn.executemany(sql, params)
        conn.commit()
        return cursor.rowcount

def update(table: str, record_id: int, data: Dict[str, Any]) -> bool:
    with write_lock():
        conn = get_conn()
        version = data.pop('version', None)
        if version is not None:
            current = conn.execute(f"SELECT version FROM {table} WHERE id=?", (record_id,)).fetchone()
            if not current or current[0] != version:
                return False
            data['version'] = version + 1
        data['updated_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        set_clause = ', '.join(f"{k}=?" for k in data)
        where = "id=? AND version=?" if version is not None else "id=?"
        params = list(data.values()) + [record_id]
        if version is not None:
            params.append(version)
        sql = f"UPDATE {table} SET {set_clause} WHERE {where}"
        conn.execute(sql, params)
        conn.commit()
        return True

def delete(table: str, record_id: int) -> bool:
    with write_lock():
        conn = get_conn()
        conn.execute(f"DELETE FROM {table} WHERE id=?", (record_id,))
        conn.commit()
        return True

def delete_where(table: str, **conditions) -> int:
    with write_lock():
        conn = get_conn()
        if not conditions:
            cursor = conn.execute(f"DELETE FROM {table}")
        else:
            where_clause = ' AND '.join(f"{k}=?" for k in conditions)
            params = list(conditions.values())
            cursor = conn.execute(f"DELETE FROM {table} WHERE {where_clause}", params)
        conn.commit()
        return cursor.rowcount

def query(table: str, columns: str = "*", where: str = "",
          params: tuple = (), order_by: str = "", limit: int = 0,
          offset: int = 0, as_dict: bool = True) -> List[Any]:
    conn = get_conn()
    sql = f"SELECT {columns} FROM {table}"
    if where:
        sql += f" WHERE {where}"
    if order_by:
        sql += f" ORDER BY {order_by}"
    if limit:
        sql += f" LIMIT {limit}"
    if offset:
        sql += f" OFFSET {offset}"
    rows = conn.execute(sql, params).fetchall()
    return [dict(r) for r in rows] if as_dict else rows

def count(table: str, where: str = "", params: tuple = ()) -> int:
    conn = get_conn()
    sql = f"SELECT COUNT(*) FROM {table}"
    if where:
        sql += f" WHERE {where}"
    r = conn.execute(sql, params).fetchone()
    return r[0] if r else 0

def raw(sql: str, params: tuple = (), fetch: bool = True) -> Any:
    conn = get_conn()
    if fetch:
        return [dict(r) for r in conn.execute(sql, params).fetchall()]
    else:
        with write_lock():
            cursor = conn.execute(sql, params)
            conn.commit()
            return cursor.rowcount

def generate_batch_id() -> str:
    return f"BATCH_{datetime.now().strftime('%Y%m%d%H%M%S')}_{os.urandom(4).hex()}"

if __name__ == "__main__":
    init_db()
    # 测试加密
    test = encrypt_password("mypassword")
    print(f"加密: {test}")
    print(f"解密: {decrypt_password(test)}")
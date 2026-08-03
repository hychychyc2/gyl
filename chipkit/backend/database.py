"""
芯片齐套管理系统 - 数据库层
SQLite WAL模式 + 应用层乐观锁
"""
import sqlite3
import threading
import time
import json
import os
from datetime import datetime
from contextlib import contextmanager
from typing import Optional, List, Dict, Any, Tuple

DB_PATH = os.path.join(os.path.dirname(__file__), "..", "data", "chipkit.db")

# 写入锁：保证同一时刻只有一个写入
_write_lock = threading.Lock()
# 连接池：每个线程一个连接
_local = threading.local()

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
    """获取写锁，超时5秒"""
    acquired = _write_lock.acquire(timeout=5)
    if not acquired:
        raise RuntimeError("写入锁获取超时，系统繁忙")
    try:
        yield
    finally:
        _write_lock.release()

def init_db():
    """初始化所有表"""
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = get_conn()
    cursor = conn.cursor()

    tables = [
        # 出货明细
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
            source TEXT DEFAULT 'manual',
            import_batch TEXT DEFAULT '',
            version INTEGER DEFAULT 1,
            created_at TEXT DEFAULT (datetime('now','localtime')),
            updated_at TEXT DEFAULT (datetime('now','localtime'))
        )""",

        # 唯一索引
        """CREATE UNIQUE INDEX IF NOT EXISTS idx_shipping_unique 
        ON shipping_detail(entity, ship_date, device_pn, wafer_lot_id, marking, 
                          good_qty, bin, invoice_no, test_program, osat, ship_to, test_wo)""",

        # 库存统一表
        """CREATE TABLE IF NOT EXISTS inventory (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            device TEXT DEFAULT '',
            marking TEXT DEFAULT '',
            qty INTEGER DEFAULT 0,
            bin TEXT DEFAULT '',
            test_program TEXT DEFAULT '',
            location_code TEXT DEFAULT '',
            warehouse_type TEXT DEFAULT '',
            warehouse_name TEXT DEFAULT '',
            batch TEXT DEFAULT '',
            date_code TEXT DEFAULT '',
            material_code TEXT DEFAULT '',
            status TEXT DEFAULT '正常',
            remark TEXT DEFAULT '',
            device_prog_bin TEXT DEFAULT '',
            import_batch TEXT DEFAULT '',
            version INTEGER DEFAULT 1,
            created_at TEXT DEFAULT (datetime('now','localtime')),
            updated_at TEXT DEFAULT (datetime('now','localtime'))
        )""",

        # 机型对照表
        """CREATE TABLE IF NOT EXISTS model_mapping (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
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
            exclusive_bin INTEGER DEFAULT 0,
            product TEXT DEFAULT '',
            osat_model TEXT DEFAULT '',
            chip_total INTEGER DEFAULT 0,
            unit_count INTEGER DEFAULT 0,
            project TEXT DEFAULT '',
            version INTEGER DEFAULT 1,
            created_at TEXT DEFAULT (datetime('now','localtime')),
            updated_at TEXT DEFAULT (datetime('now','localtime'))
        )""",
        """CREATE UNIQUE INDEX IF NOT EXISTS idx_model_unique 
        ON model_mapping(device, test_program, bin)""",

        # 用量对照表
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

        # 混BIN分配表
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

        # 外协代码映射表
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
            created_at TEXT DEFAULT (datetime('now','localtime'))
        )""",
        """CREATE UNIQUE INDEX IF NOT EXISTS idx_sub_unique 
        ON subcontractor_mapping(internal_code, external_name)""",

        # 物流时间表
        """CREATE TABLE IF NOT EXISTS logistics_time (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            destination TEXT DEFAULT '',
            transit_days INTEGER DEFAULT 0,
            latest_ship_day TEXT DEFAULT '',
            version INTEGER DEFAULT 1,
            created_at TEXT DEFAULT (datetime('now','localtime')),
            updated_at TEXT DEFAULT (datetime('now','localtime'))
        )""",

        # 料号Device对照表
        """CREATE TABLE IF NOT EXISTS material_device (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            erp_code TEXT DEFAULT '',
            device TEXT DEFAULT '',
            wafer_pn TEXT DEFAULT '',
            description TEXT DEFAULT '',
            package_desc TEXT DEFAULT '',
            version INTEGER DEFAULT 1,
            created_at TEXT DEFAULT (datetime('now','localtime'))
        )""",

        # 齐套达成表
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

        # 出货计划表
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

        # ERP库存表
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
            created_at TEXT DEFAULT (datetime('now','localtime'))
        )""",

        # 邮件配置表
        """CREATE TABLE IF NOT EXISTS email_config (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            purpose TEXT DEFAULT '',
            description TEXT DEFAULT '',
            email_address TEXT DEFAULT '',
            imap_server TEXT DEFAULT '',
            account TEXT DEFAULT '',
            password_blob TEXT DEFAULT '',
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

        # 用户表
        """CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT UNIQUE NOT NULL,
            name TEXT DEFAULT '',
            role TEXT DEFAULT 'viewer',
            password_hash TEXT DEFAULT '',
            active INTEGER DEFAULT 1,
            created_at TEXT DEFAULT (datetime('now','localtime'))
        )""",

        # 操作日志表
        """CREATE TABLE IF NOT EXISTS operation_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            action TEXT DEFAULT '',
            table_name TEXT DEFAULT '',
            record_id INTEGER DEFAULT 0,
            detail TEXT DEFAULT '',
            operator TEXT DEFAULT '',
            created_at TEXT DEFAULT (datetime('now','localtime'))
        )""",

        # 导入批次跟踪
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
            print(f"SQL error: {e}\n{sql[:200]}")

    conn.commit()

    # 创建默认管理员
    try:
        cursor.execute("SELECT id FROM users WHERE email='yuchuan.he@casue.com'")
        if not cursor.fetchone():
            cursor.execute(
                "INSERT INTO users (email, name, role, password_hash) VALUES (?,?,?,?)",
                ("yuchuan.he@casue.com", "何宇川", "admin", "chipkit_admin_2026")
            )
            conn.commit()
    except:
        pass

    print("✅ 数据库初始化完成")

def insert(table: str, data: Dict[str, Any]) -> int:
    """插入一条记录，返回id"""
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
    """批量插入，返回插入行数"""
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
    """乐观锁更新"""
    with write_lock():
        conn = get_conn()
        version = data.pop('version', None)
        if version is not None:
            current = conn.execute(f"SELECT version FROM {table} WHERE id=?", (record_id,)).fetchone()
            if not current or current[0] != version:
                return False
            new_version = version + 1
            data['version'] = new_version
        data['updated_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        set_clause = ', '.join(f"{k}=?" for k in data)
        where = "id=? AND version=?" if version is not None else "id=?"
        params = list(data.values())
        params.append(record_id)
        if version is not None:
            params.append(version)
        sql = f"UPDATE {table} SET {set_clause} WHERE {where}"
        cursor = conn.execute(sql, params)
        conn.commit()
        return cursor.rowcount > 0

def delete(table: str, record_id: int) -> bool:
    """删除记录"""
    with write_lock():
        conn = get_conn()
        cursor = conn.execute(f"DELETE FROM {table} WHERE id=?", (record_id,))
        conn.commit()
        return cursor.rowcount > 0

def delete_where(table: str, **conditions) -> int:
    """条件删除"""
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
    """通用查询"""
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
    if as_dict:
        return [dict(r) for r in rows]
    return rows

def count(table: str, where: str = "", params: tuple = ()) -> int:
    """计数查询"""
    conn = get_conn()
    sql = f"SELECT COUNT(*) FROM {table}"
    if where:
        sql += f" WHERE {where}"
    r = conn.execute(sql, params).fetchone()
    return r[0] if r else 0

def raw(sql: str, params: tuple = (), fetch: bool = True) -> Any:
    """执行原始SQL"""
    conn = get_conn()
    if fetch:
        rows = conn.execute(sql, params).fetchall()
        return [dict(r) for r in rows]
    else:
        with write_lock():
            cursor = conn.execute(sql, params)
            conn.commit()
            return cursor.rowcount

def generate_batch_id() -> str:
    return f"BATCH_{datetime.now().strftime('%Y%m%d%H%M%S')}_{os.urandom(4).hex()}"

# 清理旧连接
def close_connections():
    if hasattr(_local, 'conn') and _local.conn:
        try:
            _local.conn.close()
        except:
            pass
        _local.conn = None

if __name__ == "__main__":
    init_db()
    print("Tables created.")

# 🦞 芯片齐套管理系统 v2.0

将芯片齐套表从 Excel 升级为 **模块化 Web 管理系统**，支持库存管理、出货明细、机型对照、混BIN分配、齐套达成追踪、出货计划自动生成。

## 📊 功能模块

| 模块 | 功能 |
|---|---|
| 📊 仪表盘 | 数据总览、库存分布、最近出货 |
| 📦 库存总览 | 多仓库合并（OSAT/保税仓/其他仓/Hold/EMS），搜索筛选，在线编辑 |
| 🚚 出货明细 | 出货记录查询，过期明细高亮 |
| 🔗 机型对照 | 芯片→BIN→机型映射，关联库存显示可做台数 |
| 🔀 混BIN分配 | 多BIN组合管理，库存折算 |
| ✅ 齐套达成 | 各外协齐套追踪，欠料自动计算 |
| 📋 出货计划 | 按优先级（OSAT>保税仓>其他仓）自动生成 |
| 🏭 ERP库存 | 导入ERP库存格式 |
| 🗺️ 映射管理 | 外协代码/物流时间/料号Device 在线增删改 |
| 📧 邮件配置 | 配置邮件采集规则，支持多邮箱 |
| 📥 数据导入 | JSON 数据导入/导出 |

## 🚀 Windows 部署

### 前置条件
- Python 3.9+
- 依赖包：`openpyxl`、`xlrd`

### 1. 安装依赖

```powershell
pip install openpyxl xlrd
```

### 2. 数据迁移（从 Excel 导入）

```powershell
cd chipkit\scripts
set PYTHONPATH=..\backend;%PYTHONPATH%
python migrate_fast.py
python migrate_inventory.py
```

> 需要将 `芯片齐套表_lastest (78).xlsx` 放在 `chipkit` 同级目录下（即 `..\` 位置）。

### 3. 启动服务

```powershell
cd chipkit\backend
set PYTHONPATH=.
python server.py
```

### 4. 访问系统

打开浏览器访问：**`http://localhost:8765`**

## 📁 项目结构

```
chipkit/
├── backend/
│   ├── server.py          # HTTP 服务器 + API 路由
│   ├── database.py        # SQLite 数据库层（WAL模式，支持高并发）
│   └── email_collector.py # 邮件采集引擎（IMAP）
├── frontend/
│   ├── index.html         # 前端页面
│   └── app.js             # 前端逻辑（纯JS，零框架）
├── scripts/
│   ├── migrate_fast.py    # 基础数据迁移（机型对照/出货明细/用量等）
│   └── migrate_inventory.py # 库存数据迁移
├── data/                  # SQLite 数据库文件
├── exports/               # 导出文件目录
├── start.sh               # Linux 启动脚本
└── README.md
```

## 🔐 数据库设计

- **SQLite (WAL 模式)**：支持高并发读取，写入串行化
- **应用层锁**：`threading.Lock` 保证写入原子性
- **乐观锁**：version 字段防止并发冲突
- **去重**：UNIQUE 索引防止重复数据

## ⚙️ API 接口

| 接口 | 方法 | 说明 |
|---|---|---|
| `/api/dashboard` | GET | 仪表盘数据 |
| `/api/query/{table}` | POST | 通用查询 |
| `/api/insert/{table}` | POST | 插入记录 |
| `/api/update/{table}/{id}` | PUT | 更新记录 |
| `/api/delete/{table}/{id}` | DELETE | 删除记录 |
| `/api/insert_many/{table}` | POST | 批量插入 |
| `/api/inventory/summary` | GET | 库存汇总 |
| `/api/inventory/with_model` | GET | 库存关联机型 |
| `/api/shipping/expired` | GET | 过期出货明细 |
| `/api/model/mapping` | GET | 机型对照 |
| `/api/model/with_stock` | GET | 机型关联库存 |
| `/api/kit/completion` | GET | 齐套达成 |
| `/api/kit/calculate_shortage` | POST | 计算欠料 |
| `/api/shipping/auto_plan` | POST | 自动生成出货计划 |
| `/api/upload/inventory` | POST | 上传库存文件 |
| `/api/upload/shipping` | POST | 上传出货明细 |
| `/api/upload/erp_inventory` | POST | 上传ERP库存 |
| `/api/export/{table}` | GET | 导出数据为JSON |

## 📧 邮件采集配置

在"📧 邮件配置"模块中，可以添加邮件采集规则，支持以下用途：

- `shipping_detail`：出货明细（OSAT shipping list）
- `inventory`：库存报表
- `model_mapping`：机型对照表
- `mix_bin`：混BIN关系
- `order_allocation`：订单分配（张胜文邮件）
- `hold_inventory`：Hold库存

## 🛠️ 技术栈

- **后端**：Python 3 + 内置 http.server（零外部依赖）
- **数据库**：SQLite (WAL)
- **前端**：HTML5 + Vanilla JS（零框架，响应式设计）
- **双模式**：自动检测后端 API，离线时切换 IndexedDB 本地存储

## 📝 后续开发计划

- [ ] PO Agent 集成：匹配采购PO到出货明细
- [ ] 销售出库模板生成（保留VBA宏）
- [ ] 邮件自动定时采集（OpenClaw cron）
- [ ] 齐套达成各月大计划自动抓取
- [ ] 多用户权限管理
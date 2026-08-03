# 🦞 芯片齐套管理系统 v2.0

将芯片齐套表从 Excel 升级为 **模块化 Web 管理系统**。

## 功能模块

| 模块 | 功能 | 数据来源 |
|---|---|---|
| 📊 仪表盘 | 数据总览、库存分布、一键采集 | 汇总 |
| 📦 库存总览 | 多仓库合并展示、搜索筛选、在线编辑 | 邮件自动 + 手动上传 |
| 📊 库存透视 | 按机型查看各仓库芯片数量和可做台数 | 库存 + 机型 + 用量联动 |
| 🚚 出货明细 | 出货记录查询、过期明细高亮 | 邮件自动 + 手动上传 |
| 🔗 机型对照 | 芯片→BIN→机型映射、关联库存 | 邮件自动 |
| 🔀 混BIN分配 | 多BIN组合管理 | 邮件自动 |
| ✅ 齐套达成 | 各外协齐套追踪、自动欠料计算 | 邮件自动 |
| 📋 出货计划 | 按优先级自动生成 | 库存 + 齐套联动 |
| 🏭 ERP库存 | 按批次管理 | 手动上传 |
| 🗺️ 映射管理 | 外协代码/物流时间/料号Device 在线管理 | 手动 |
| 📧 邮件配置 | 邮箱采集规则配置、手动触发采集 | 手动 |
| 📤 数据导入 | MES/ERP/EMS格式Excel上传 | 手动 |

## 数据来源分类

| 类型 | 来源 | 方式 |
|---|---|---|
| 出货明细 | 邮件（各OSAT shipping list） | 自动采集 |
| OSAT库存 | 邮件（OSAT库存报表） | 自动采集 |
| Hold库存 | 邮件（Hold报表） | 自动采集 |
| 机型对照表 | 邮件（wenzhe.liu01_w@casue.com） | 自动采集 |
| 混BIN关系 | 邮件（wenzhe.liu01_w@casue.com） | 自动采集 |
| 订单分配 | 邮件（shengwen.zhang@casue.com） | 自动采集 |
| 保税仓（QHBS） | ERP格式Excel上传 | 手动 |
| 其他仓（SZKXYCL等） | MES格式Excel上传 | 手动 |
| EMS库存 | EMS格式Excel上传 | 手动 |
| 映射表 | 前端在线管理 | 手动 |

## Windows 部署

### 前置条件
- Python 3.9+
- pip

### 1. 安装依赖
```powershell
pip install openpyxl xlrd
```

> 只需这两个包，其余全是 Python 标准库（http.server、sqlite3、imaplib、email 等）。

### 2. 确保 Excel 文件
将 `芯片齐套表_lastest (78).xlsx` 放在 `chipkit/` 同级目录（即项目根目录）。

### 3. 数据迁移
```powershell
cd chipkit\scripts
$env:PYTHONPATH = "..\backend"
python migrate_fast.py
python migrate_inventory.py
```

### 4. 启动服务
```powershell
cd ..\backend
$env:PYTHONPATH = "."
python server.py
```

### 5. 访问
浏览器打开 **http://localhost:8765**

## 定时采集

服务启动后自动开启定时任务：
- **每天 9:00** 同步采集所有活跃邮箱配置
- **每天 21:00** 同步采集所有活跃邮箱配置
- 也可在仪表盘点击"手动采集全部"或通过邮件配置模块单条触发

## 邮件配置

在"📧 邮件配置"模块添加邮箱规则，支持以下用途：

| 用途 | 说明 |
|---|---|
| shipping_detail | 出货明细 |
| osat_inventory | OSAT库存 |
| hold_inventory | Hold库存 |
| model_mapping | 机型对照表 |
| mix_bin | 混BIN关系 |
| order_allocation | 订单分配 |

每个 OSAT 的出货明细和库存需要分别添加独立规则。

邮箱密码使用 XOR + Base64 加密存储，不存明文。

## 库存导入

在"📤 数据导入"模块支持三种格式：

**MES 格式**（SZKXYCL/HSJXYCL 等）：
- 自动解析"生产日期/批次"字段中的 marking/bin/程序
- 标记为 `warehouse_type=other`

**ERP 格式**（QHBS 等）：
- 自动解析"批次"字段中的 marking/bin/程序
- 标记为 `warehouse_type=bonded`

**EMS 格式**（芯片结存统计）：
- 读取"各外EMS外协库存明细"sheet
- 标记为 `warehouse_type=ems`

上传自动覆盖对应仓库类型+名称的旧数据。

## 技术栈

- **后端**：Python 3 + 内置 http.server（零外部依赖）
- **数据库**：SQLite (WAL模式 + 应用层写锁 + 乐观锁)
- **前端**：HTML5 + Vanilla JS（零框架，响应式）
- **邮件**：imaplib（Python内置）
- **Excel**：openpyxl + xlrd
- **密码**：XOR + Base64 加密存储

## 项目结构

```
chipkit/
├── backend/
│   ├── server.py          # HTTP服务器 + API + 定时任务
│   ├── database.py        # SQLite数据库层 + 密码加密
│   └── email_collector.py # 邮件采集引擎
├── frontend/
│   ├── index.html         # 前端页面
│   └── app.js             # 前端逻辑
├── scripts/
│   ├── migrate_fast.py    # 基础数据迁移
│   └── migrate_inventory.py # 库存数据迁移
├── data/                  # SQLite数据库
├── exports/               # 导出目录
└── README.md
```
# Excel邮件附件自动处理工具

## 功能介绍
该工具可自动从指定邮箱文件夹下载最新匹配的Excel附件，智能兼容`.xls`/`.xlsx`格式（自动检测文件真实格式并修正扩展名），按自定义条件筛选数据后，写入指定Excel文件（支持覆盖/追加模式，可配置数据去重）。

核心特性：
- 📧 支持IMAP协议邮箱的附件下载
- 📄 智能兼容Excel格式（自动解决扩展名与真实格式不符问题）
- ⚡ 自定义数据筛选规则
- 📝 支持数据覆盖/追加写入，可配置重复数据删除
- 📜 详细运行日志，便于问题排查

## 环境准备

### 1. Python版本要求
Python 3.7+

### 2. 安装依赖包
```bash
pip install openpyxl xlrd imaplib email-validator
```

> 注意：xlrd推荐版本2.0.1（兼容.xls格式），openpyxl推荐3.0+（兼容.xlsx格式）

## 配置文件说明
工具核心配置文件为`config.json`，需放在脚本同级目录，完整字段说明如下：

### 整体结构
```json
{
  "email_config": {},
  "rules": [],
  "other": {}
}
```

### 1. email_config（邮箱配置，必填）
| 字段名 | 类型 | 说明 | 示例 |
|--------|------|------|------|
| imap_server | string | 邮箱IMAP服务器地址 | 网易邮箱：`imap.163.com`，QQ邮箱：`imap.qq.com`，企业微信邮箱：`imap.weixin.qq.com` |
| account | string | 邮箱账号 | `your_email@163.com` |
| password | string | 邮箱授权码（非登录密码） | 需在邮箱设置中开启IMAP并生成授权码 |
| root_folder | string | 邮件根文件夹 | 默认为`INBOX`（收件箱） |
| exclude_folders | array | 排除的子文件夹 | `["已删除", "垃圾邮件"]` |
| search_criteria | string | 邮件搜索条件（IMAP格式） | `SINCE "24-Dec-2025" BEFORE "25-Dec-2025"`（可选，留空则默认取当日邮件） |

> 授权码获取说明：
> - 163邮箱：设置 → POP3/SMTP/IMAP → 开启IMAP → 生成授权码
> - QQ邮箱：设置 → 账户 → POP3/IMAP/SMTP/Exchange/CardDAV/CalDAV服务 → 开启IMAP/SMTP服务 → 生成授权码

### 2. rules（处理规则，必填，支持多规则）
数组类型，每个元素对应一个附件处理规则，字段说明：

#### 2.1 attach_rule（附件解析规则）
| 字段名 | 类型 | 说明 | 示例 |
|--------|------|------|------|
| match_key | string | 附件名称匹配关键词（模糊匹配） | `RPT_STOCK_JSCC` |
| suffix | string | 附件后缀 | `.xls`（支持.xls/.xlsx，工具会自动兼容真实格式） |
| sheet | string | 附件中要解析的Sheet名称 | `库存数据` |
| header_row | int | 表头行号（从1开始） | `1` |
| cols_range | object | 要提取的列范围 | `{"start_col": "A", "end_col": "D"}`（提取A-D列） |
| filter_conditions | array | 数据筛选条件（可选） | 见下方筛选条件示例 |

#### 2.2 local_rule（本地写入规则）
| 字段名 | 类型 | 说明 | 示例 |
|--------|------|------|------|
| file_path | string | 本地目标Excel文件路径 | `E:\供应链\芯片\总表\库存汇总.xlsx` |
| sheet | string | 要写入的Sheet名称 | `JSCC库存` |
| header_row | int | 目标文件表头行号（从1开始） | `1` |
| cols_range | object | 要写入的列范围 | `{"start_col": "B", "end_col": "E"}`（写入B-E列） |
| write_mode | string | 写入模式 | `overwrite`（覆盖）/ `append`（追加） |
| deduplicate | bool | 是否开启去重 | `true`/`false` |
| deduplicate_cols | array | 去重依据列（开启去重时必填） | `["物料编码", "日期"]`（按物料编码+日期去重） |

#### 筛选条件示例
```json
"filter_conditions": [
  {
    "col": "库存数量",  // 要筛选的列名（对应表头）
    "operator": "gt",  // 操作符：eq(等于)/ne(不等于)/gt(大于)/lt(小于)/ge(大于等于)/le(小于等于)/contains(包含)/not_contains(不包含)/between(区间)/in(包含于)
    "value": 100       // 筛选值：between填数组如[100, 500]，in填数组如["A类", "B类"]
  },
  {
    "col": "物料类型",
    "operator": "in",
    "value": ["芯片", "电阻"]
  }
]
```

### 3. other（其他配置）
| 字段名 | 类型 | 说明 | 示例 |
|--------|------|------|------|
| temp_dir | string | 临时文件目录（存放下载的附件） | `E:\供应链\芯片\总表\temp_attachments` |
| clean_temp | bool | 运行结束后是否清理临时目录 | `true` |

### 完整配置示例
```json
{
  "email_config": {
    "imap_server": "imap.163.com",
    "account": "your_email@163.com",
    "password": "your_auth_code",
    "root_folder": "INBOX",
    "exclude_folders": ["已删除", "垃圾邮件"],
    "search_criteria": "SINCE \"24-Dec-2025\" BEFORE \"25-Dec-2025\""
  },
  "rules": [
    {
      "attach_rule": {
        "match_key": "RPT_STOCK_JSCC",
        "suffix": ".xls",
        "sheet": "库存数据",
        "header_row": 1,
        "cols_range": {
          "start_col": "A",
          "end_col": "D"
        },
        "filter_conditions": [
          {
            "col": "库存数量",
            "operator": "gt",
            "value": 100
          }
        ]
      },
      "local_rule": {
        "file_path": "E:\\供应链\\芯片\\总表\\库存汇总.xlsx",
        "sheet": "JSCC库存",
        "header_row": 1,
        "cols_range": {
          "start_col": "B",
          "end_col": "E"
        },
        "write_mode": "append",
        "deduplicate": true,
        "deduplicate_cols": ["物料编码"]
      }
    }
  ],
  "other": {
    "temp_dir": "E:\\供应链\\芯片\\总表\\temp_attachments",
    "clean_temp": true
  }
}
```

## 使用步骤

### 1. 配置config.json
根据上述字段说明，填写`config.json`文件，重点确认：
- 邮箱IMAP服务器、账号、授权码正确
- 附件匹配关键词、Sheet名称、列范围准确
- 目标文件路径存在且有读写权限

### 2. 运行脚本
```bash
python your_script_name.py
```

### 3. 查看运行日志
脚本会输出详细运行日志，包括：
- 邮箱文件夹解析结果
- 附件下载状态
- 文件格式检测与重命名（如有）
- 数据提取与筛选结果
- 本地文件写入状态
- 最终处理结果

## 常见问题

### Q1: 提示“读取文件彻底失败”
- 原因：文件扩展名与真实格式不符，且自动重命名后仍无法读取
- 解决：手动将附件改为正确扩展名（如.xls改为.xlsx），或检查文件是否损坏

### Q2: 邮箱登录失败
- 原因：使用了登录密码而非授权码，或IMAP服务未开启
- 解决：在邮箱设置中开启IMAP，生成并使用授权码

### Q3: 提示“未找到列：XXX”
- 原因：筛选条件中的列名与表头不匹配，或表头行号设置错误
- 解决：核对附件表头名称，确认`header_row`字段正确

### Q4: 数据写入后乱码
- 原因：文件编码不兼容
- 解决：将目标Excel文件另存为.xlsx格式，或检查附件编码

### Q5: 去重功能未生效
- 原因：去重列名与目标文件表头不匹配，或`deduplicate`未设为`true`
- 解决：核对去重列名，确认`deduplicate`为`true`

## 注意事项
1. 确保目标Excel文件未被其他程序占用（如Excel客户端打开），否则会写入失败
2. 临时目录需有读写权限，`clean_temp`设为`true`时，运行结束后会自动删除临时文件
3. 筛选条件中的数值类型需与附件中数据类型一致（如数字不要加引号）
4. 多规则配置时，确保每个规则的附件关键词不重复，避免下载错误附件
5. 建议定期备份目标Excel文件，避免覆盖/追加操作导致数据丢失
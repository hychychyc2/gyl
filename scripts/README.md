# 采购订单自动化脚本使用指南

## 功能概述

本脚本每天自动完成以下任务：

1. ✅ 从IMAP读取 `na.yang_w@casue.com` 当日邮件
2. ✅ 解析邮件提取芯片型号、数量、主体等信息
3. ✅ 从价格表匹配芯片价格
4. ✅ 生成采购订单号（如 SZK202605110001）
5. ✅ 填充Oracle WebADI导入模板
6. ⏳ 操作Oracle客户端导入（需手动或配置自动化）
7. ✅ 更新统计表格
8. ✅ 发送结果邮件报告

---

## 安装步骤

### 1. 安装Python

从 https://www.python.org/downloads/ 下载Python 3.8+

安装时勾选：
- ✅ Add Python to PATH
- ✅ Install pip

### 2. 下载脚本和依赖

将以下文件夹复制到你的电脑：
```
scripts/
data/
config/
```

在 `scripts` 目录下运行：
```cmd
pip install -r requirements.txt
```

### 3. 验证安装

```cmd
python purchase_order_automation.py
```

---

## 使用方法

### 手动运行

双击 `run_automation.bat` 或在命令行运行：
```cmd
python purchase_order_automation.py
```

### 定时运行（Windows任务计划）

#### 方法1：使用Windows任务计划程序

1. 打开 **任务计划程序**（搜索"Task Scheduler"）
2. 点击 **创建基本任务**
3. 名称：`采购订单自动化`
4. 触发器：**每天 22:00**
5. 操作：**启动程序**
   - 程序：`python.exe` 的完整路径（如 `C:\Python310\python.exe`）
   - 参数：`purchase_order_automation.py`
   - 起始于：脚本目录路径

#### 方法2：使用批处理文件

1. 创建基本任务
2. 触发器：**每天 22:00**
3. 操作：**启动程序**
   - 程序：`run_automation.bat` 的完整路径

---

## Oracle自动化配置

脚本支持自动操作Oracle客户端，需要额外配置：

### 依赖安装

```cmd
pip install pyautogui pyperclip
```

### 使用方法

1. 确保Oracle客户端在脚本运行前已打开
2. 脚本会自动定位Oracle窗口并执行导入操作
3. 可能需要调整坐标位置（根据你的Oracle界面）

### 坐标调整

编辑 `purchase_order_automation.py` 中的 `automate_oracle_import` 函数：
```python
# 根据你的Oracle界面调整坐标
pyautogui.click(x=100, y=100)  # 导入按钮位置
```

**如何获取坐标**：
```python
import pyautogui
print(pyautogui.position())  # 鼠标当前位置
```

---

## 文件结构

```
workspace-schedule_tasks/
├── config/
│   ├── credentials.enc    # 加密凭证文件
│   └── .key               # 解密密钥
├── data/
│   ├── templates/
│   │   └── webadi_template.xlsm    # Oracle导入模板
│   ├── statistics/
│   │   ├── domestic_statistics.xlsx    # 国内统计表
│   │   └── international_statistics.xlsx # 国际统计表
│   ├── prices/
│   │   └── current_prices.xlsx    # 价格表（每月更新）
│   └── output/
│       └── SZK202605110001.xlsm   # 生成的订单文件
├── scripts/
│   ├── purchase_order_automation.py   # 主脚本
│   ├── requirements.txt               # 依赖
│   ├── run_automation.bat             # Windows启动器
│   └── README.md                      # 本说明文件
└── memory/
    └── 2026-05-11.md                  # 日志文件
```

---

## 主体识别规则

脚本自动识别邮件中的主体：

| 主体代码 | 名称 | 币种 | 订单号前缀 |
|---------|------|------|-----------|
| SZK | 世纪云芯 | CNY | SZK |
| ICK | 智能云芯 | CNY | ICK |
| HSJ | 海南世纪 | CNY | HSJ |
| DPT | Bitmain Development PTE. LTD. | USD | DPT |
| BJK | Bitmain Beijing | CNY | BJK |

---

## 价格表更新

每月更新价格表：

1. 替换 `data/prices/current_prices.xlsx`
2. 确保"PO"工作表包含：
   - 第1列：芯片型号
   - 第4列：USD价格（保留四位小数）

---

## 常见问题

### Q: 邮件读取失败？

检查：
- IMAP服务器地址是否正确
- 邮箱密码是否正确
- 网络连接是否正常

### Q: 价格匹配失败？

检查：
- 价格表中的型号是否包含邮件中的型号
- 型号格式是否一致（如 BM1362AA vs BM1362）

### Q: Oracle自动化不工作？

需要：
- 手动配置界面坐标
- 确保Oracle窗口在屏幕可见区域

---

## 联系支持

如有问题，请在工作空间留言。
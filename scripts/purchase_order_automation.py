#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
采购订单自动化脚本
每天晚上10点运行，自动完成：
1. 读取杨娜邮件获取订单信息
2. 匹配价格表
3. 生成采购订单号
4. 填充WebADI模板
5. 操作Oracle导入
6. 更新统计表格
7. 发送结果邮件

使用方法：
    python purchase_order_automation.py

依赖安装：
    pip install imapclient openpyxl pyautogui pyperclip python-dateutil
"""

import os
import sys
import json
import imaplib
import email
from email.header import decode_header
from datetime import datetime, date
import re
import openpyxl
from openpyxl import Workbook, load_workbook
import time
import subprocess
from pathlib import Path

# ============== 配置区 ==============

# 凭证配置（从加密文件读取，或直接配置）
CONFIG = {
    "email": {
        "account": "yuchuan.he@casue.com",
        "password": "-DxpOD5kkN)(RuPgAK-p",
        "imap_server": "imap.appia.vip",
        "source_email": "na.yang_w@casue.com"
    },
    "erp": {
        "username": "607693",
        "password": "hyc010815"
    },
    "report_email": "yuchuan.he@casue.com",  # 报告发送目标
}

# 文件路径配置
BASE_DIR = Path(__file__).parent.parent
DATA_DIR = BASE_DIR / "data"

TEMPLATES_DIR = DATA_DIR / "templates"
STATISTICS_DIR = DATA_DIR / "statistics"
PRICES_DIR = DATA_DIR / "prices"
OUTPUT_DIR = DATA_DIR / "output"

# 主体配置
ENTITY_CONFIG = {
    "SZK": {"name": "世纪云芯", "currency": "CNY", "prefix": "SZK"},
    "ICK": {"name": "智能云芯", "currency": "CNY", "prefix": "ICK"},
    "HSJ": {"name": "海南世纪", "currency": "CNY", "prefix": "HSJ"},
    "DPT": {"name": "Bitmain Development PTE. LTD.", "currency": "USD", "prefix": "DPT"},
    "BJK": {"name": "Bitmain Beijing", "currency": "CNY", "prefix": "BJK"},
}

# ============== 邮件读取模块 ==============

def decode_email_header(header):
    """解码邮件头部"""
    if header is None:
        return ""
    decoded_parts = decode_header(header)
    result = []
    for part, charset in decoded_parts:
        if isinstance(part, bytes):
            if charset:
                try:
                    result.append(part.decode(charset))
                except:
                    result.append(part.decode('utf-8', errors='ignore'))
            else:
                result.append(part.decode('utf-8', errors='ignore'))
        else:
            result.append(str(part))
    return ''.join(result)

def get_today_emails_from_source():
    """获取今日来自杨娜的邮件"""
    print("正在连接邮件服务器...")
    
    imap = imaplib.IMAP4_SSL(CONFIG["email"]["imap_server"])
    imap.login(CONFIG["email"]["account"], CONFIG["email"]["password"])
    
    # 邮件在 MC/po 子文件夹，不在主收件箱
    target_folder = "MC/po"
    print(f"搜索文件夹: {target_folder}")
    
    try:
        status = imap.select(target_folder)
        if status[0] != "OK":
            print(f"无法访问文件夹 {target_folder}，尝试搜索收件箱")
            imap.select("INBOX")
    except Exception as e:
        print(f"文件夹访问失败: {e}，使用收件箱")
        imap.select("INBOX")
    
    # 搜索今日邮件
    today = date.today().strftime("%d-%b-%Y")
    search_criteria = f'(FROM "{CONFIG["email"]["source_email"]}" ON "{today}")'
    
    status, messages = imap.search(None, search_criteria)
    
    if status != "OK":
        print("搜索邮件失败")
        imap.logout()
        return []
    
    email_ids = messages[0].split()
    print(f"找到 {len(email_ids)} 封今日邮件")
    
    emails_data = []
    for email_id in email_ids:
        status, msg_data = imap.fetch(email_id, "(RFC822)")
        if status == "OK":
            msg = email.message_from_bytes(msg_data[0][1])
            
            subject = decode_email_header(msg.get("Subject"))
            from_addr = decode_email_header(msg.get("From"))
            date_str = decode_email_header(msg.get("Date"))
            
            # 提取邮件正文
            body = ""
            if msg.is_multipart():
                for part in msg.walk():
                    content_type = part.get_content_type()
                    if content_type == "text/plain" or content_type == "text/html":
                        try:
                            charset = part.get_content_charset() or 'utf-8'
                            payload = part.get_payload(decode=True)
                            if payload:
                                body += payload.decode(charset, errors='ignore')
                        except:
                            pass
            else:
                try:
                    charset = msg.get_content_charset() or 'utf-8'
                    payload = msg.get_payload(decode=True)
                    if payload:
                        body = payload.decode(charset, errors='ignore')
                except:
                    pass
            
            emails_data.append({
                "id": email_id.decode(),
                "subject": subject,
                "from": from_addr,
                "date": date_str,
                "body": body
            })
    
    imap.logout()
    return emails_data

def parse_order_info_from_email(email_body, subject):
    """从邮件内容解析订单信息"""
    print(f"解析邮件: {subject}")
    
    order_info = {
        "items": [],  # 芯片型号和数量列表
        "entity": None,  # 主体
        "address": None,  # 地址
        "test_factory": None,  # 测试厂
        "notes": None  # 备注
    }
    
    # 尝试解析芯片型号和数量
    # 常见格式: BM1362 10000pcs 或 BM1362AC: 5000等
    
    # 模式1: 型号 + 数量
    pattern1 = r'(BM\d{4}[A-Z]{0,3}[\+\w]*)\s*[:\s]\s*(\d+)\s*(pcs|个|片)?'
    matches1 = re.findall(pattern1, email_body, re.IGNORECASE)
    
    # 模式2: 型号 + 数量（更宽松）
    pattern2 = r'(BM\d{4}[A-Z]{0,3}[\+\w]*)\s+(\d+)'
    matches2 = re.findall(pattern2, email_body, re.IGNORECASE)
    
    # 模式3: 从表格格式解析（如果有）
    pattern3 = r'(\d+)\s*(pcs|个|片)?\s*(BM\d{4}[A-Z]{0,3}[\+\w]*)'
    matches3 = re.findall(pattern3, email_body, re.IGNORECASE)
    
    all_matches = matches1 + matches2 + [(m[2], m[0]) for m in matches3]
    
    for match in all_matches:
        model = match[0].upper().strip()
        quantity = int(match[1])
        order_info["items"].append({
            "model": model,
            "quantity": quantity
        })
    
    # 解析主体（从邮件内容或主题）
    for entity_code in ENTITY_CONFIG.keys():
        if entity_code in email_body.upper() or entity_code in subject.upper():
            order_info["entity"] = entity_code
            break
    
    # 如果没找到主体，根据内容推断
    if not order_info["entity"]:
        if "深圳" in email_body or "SZ" in email_body:
            order_info["entity"] = "SZK"
        elif "新加坡" in email_body or "SG" in email_body:
            order_info["entity"] = "DPT"
        elif "北京" in email_body or "BJ" in email_body:
            order_info["entity"] = "BJK"
        elif "海南" in email_body:
            order_info["entity"] = "HSJ"
    
    # 解析地址
    address_patterns = [
        r'地址[:\s：]+([^\n]+)',
        r'收货地址[:\s：]+([^\n]+)',
        r'发货至[:\s：]+([^\n]+)',
        r'送至[:\s：]+([^\n]+)',
    ]
    for pattern in address_patterns:
        match = re.search(pattern, email_body)
        if match:
            order_info["address"] = match.group(1).strip()
            break
    
    # 解析测试厂
    test_factory_patterns = ["XJ", "捷策创", "SCK", "朗华", "Vtest", "确安", "HN"]
    for tf in test_factory_patterns:
        if tf in email_body:
            order_info["test_factory"] = tf
            break
    
    # 备注
    if "备注" in email_body or "注意" in email_body:
        match = re.search(r'(备注|注意)[:\s：]+([^\n]+)', email_body)
        if match:
            order_info["notes"] = match.group(2).strip()
    
    return order_info

# ============== 价格匹配模块 ==============

def load_price_table():
    """加载价格表"""
    price_file = PRICES_DIR / "current_prices.xlsx"
    if not price_file.exists():
        print(f"价格表文件不存在: {price_file}")
        return {}
    
    wb = load_workbook(price_file, data_only=True)
    ws = wb["PO"]
    
    prices = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] and row[3]:  # 型号和价格
            model = str(row[0]).upper().strip()
            price = float(row[3])
            prices[model] = price
    
    wb.close()
    print(f"已加载 {len(prices)} 个型号的价格")
    return prices

def match_price(model, prices):
    """匹配芯片价格"""
    model_upper = model.upper().strip()
    
    # 直接匹配
    if model_upper in prices:
        return prices[model_upper]
    
    # 去掉后缀字母匹配（如BM1362AA -> BM1362）
    base_model = re.sub(r'[A-Z]{2,3}$', '', model_upper)
    if base_model in prices:
        return prices[base_model]
    
    # BM1368+ 特殊处理
    if "1368+" in model_upper or "1368PA" in model_upper or "1368PB" in model_upper:
        return prices.get("BM1368+", prices.get("BM1368", 0))
    
    print(f"警告: 未找到型号 {model} 的价格")
    return None

# ============== 订单号生成模块 ==============

def generate_order_number(entity, existing_numbers=None):
    """生成采购订单号"""
    today = date.today()
    date_str = today.strftime("%Y%m%d")
    prefix = ENTITY_CONFIG[entity]["prefix"]
    
    # 查找今日最大序号
    if existing_numbers:
        today_prefix = f"{prefix}{date_str}"
        max_seq = 0
        for num in existing_numbers:
            if str(num).startswith(today_prefix):
                try:
                    seq = int(str(num)[-4:])
                    max_seq = max(max_seq, seq)
                except:
                    pass
        next_seq = max_seq + 1
    else:
        next_seq = 1
    
    return f"{prefix}{date_str}{next_seq:04d}"

def get_existing_order_numbers(statistics_file):
    """从统计表获取现有订单号"""
    if not statistics_file.exists():
        return []
    
    wb = load_workbook(statistics_file, data_only=True)
    ws = wb.active
    
    numbers = []
    for row in ws.iter_rows(min_row=2, max_col=9, values_only=True):
        if row[8]:  # 采购订单号列
            numbers.append(row[8])
    
    wb.close()
    return numbers

# ============== WebADI模板填充模块 ==============

def fill_webadi_template(order_data, output_file):
    """
    填充WebADI模板
    
    注意：由于xlsm包含VBA宏和自定义图片，openpyxl无法直接保存。
    这里生成一个纯xlsx数据文件，用户需要：
    1. 打开WebADI模板
    2. 复制xlsx文件中的数据到模板
    3. 或者使用xlwings/win32com在本地操作（需要Excel）
    """
    from openpyxl import Workbook
    
    wb = Workbook()
    ws = wb.active
    ws.title = "采购订单数据"
    
    entity = order_data["entity"]
    entity_config = ENTITY_CONFIG[entity]
    
    # 添加标题行
    headers = [
        "加载", "业务实体", "类型", "采购订单号", "币种", "采购员", "供应商", 
        "供应商地点", "来源子库存", "收货方", "目的子库存", "收单方", "付款方式",
        "内部申请类型", "货贷", "是否报关", "加工费报价OA单据号", "摘要", "业务模式",
        "行号", "行类型", "物料", "物料说明", "单位", "数量", "创建日期", 
        "承诺日期", "需求日期", "不含税单价", "含税单价", "税率", "品牌/厂商"
    ]
    for col, header in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=header)
    
    # 填充数据行
    row_num = 2
    line_num = 1
    
    for item in order_data["items"]:
        model = item["model"]
        quantity = item["quantity"]
        price = item["price"]
        
        if price is None:
            print(f"跳过无价格的型号: {model}")
            continue
        
        # 头部信息
        ws.cell(row=row_num, column=1, value="Y")  # 加载标记
        ws.cell(row=row_num, column=2, value=entity)  # 业务实体
        ws.cell(row=row_num, column=3, value="标准采购订单")  # 类型
        ws.cell(row=row_num, column=4, value=order_data["order_number"])  # 采购订单号
        ws.cell(row=row_num, column=5, value=entity_config["currency"])  # 币种
        ws.cell(row=row_num, column=6, value="何宇川,")  # 采购员
        ws.cell(row=row_num, column=7, value="BITMAIN DEVELOPMENT PTE. LTD.")  # 供应商
        ws.cell(row=row_num, column=8, value="SG")  # 供应商地点
        ws.cell(row=row_num, column=9, value="SZKXYCL")  # 来源子库存
        ws.cell(row=row_num, column=10, value="1004.Bitmain Shenzhen")  # 收货方
        ws.cell(row=row_num, column=11, value="SZKXYCL")  # 目的子库存
        ws.cell(row=row_num, column=12, value="1004.Bitmain Shenzhen")  # 收单方
        ws.cell(row=row_num, column=13, value="付款方式一")  # 付款方式
        ws.cell(row=row_num, column=16, value="Y")  # 是否报关
        
        # 行信息
        ws.cell(row=row_num, column=20, value=line_num)  # 行号
        ws.cell(row=row_num, column=21, value="BM系列")  # 行类型
        ws.cell(row=row_num, column=22, value="Y31010544")  # 物料编码（需要根据型号调整）
        ws.cell(row=row_num, column=23, value=f"{model}芯片")  # 物料说明
        ws.cell(row=row_num, column=24, value="个")  # 单位
        ws.cell(row=row_num, column=25, value=quantity)  # 数量
        ws.cell(row=row_num, column=26, value=date.today())  # 创建日期
        ws.cell(row=row_num, column=27, value=date.today())  # 承诺日期
        ws.cell(row=row_num, column=28, value=date.today())  # 需求日期
        ws.cell(row=row_num, column=29, value=price)  # 不含税单价
        ws.cell(row=row_num, column=32, value="ANTMINER")  # 品牌/厂商
        
        row_num += 1
        line_num += 1
    
    output_file.parent.mkdir(parents=True, exist_ok=True)
    
    # 保存为xlsx格式
    xlsx_file = output_file.with_suffix('.xlsx')
    wb.save(xlsx_file)
    wb.close()
    
    print(f"数据文件已生成: {xlsx_file}")
    print("注意: 请打开WebADI模板，将此文件数据复制粘贴到模板中")
    return True

# ============== 统计表格更新模块 ==============

def update_statistics_table(order_data, entity):
    """更新统计表格"""
    if entity in ["SZK", "ICK", "HSJ", "BJK"]:
        # 国内表格
        stats_file = STATISTICS_DIR / "domestic_statistics.xlsx"
    else:
        # 国际表格
        stats_file = STATISTICS_DIR / "international_statistics.xlsx"
    
    if not stats_file.exists():
        print(f"统计文件不存在: {stats_file}")
        return False
    
    wb = load_workbook(stats_file)
    
    # 选择正确的sheet（当月）
    today = date.today()
    month_name = today.strftime("%m月")
    
    if month_name in wb.sheetnames:
        ws = wb[month_name]
    else:
        # 使用第一个sheet或创建新sheet
        ws = wb.active
    
    # 找到最后一行
    last_row = ws.max_row + 1
    
    # 添加订单记录 - 按 Excel 格式：序号, 抬头, 出货日期, 主体, 测试厂, 收货地址, 型号, 物料编码, 数量, 新PO, SO, 单价, 是否已出, 对比, 标记
    for i, item in enumerate(order_data["items"]):
        ws.cell(row=last_row + i, column=1, value=last_row + i - 1)  # 序号
        ws.cell(row=last_row + i, column=2, value="chanhua")  # 抬头
        ws.cell(row=last_row + i, column=3, value=date.today())  # 出货日期
        ws.cell(row=last_row + i, column=4, value=ENTITY_CONFIG.get(entity, {}).get("name", entity))  # 主体
        ws.cell(row=last_row + i, column=5, value=order_data.get("test_factory", "XJ"))  # 测试厂
        ws.cell(row=last_row + i, column=6, value=order_data.get("address", ""))  # 收货地址
        ws.cell(row=last_row + i, column=7, value=item["model"])  # 型号
        ws.cell(row=last_row + i, column=8, value=f"Y09{item['model']}")  # 物料编码
        ws.cell(row=last_row + i, column=9, value=item["quantity"])  # 数量
        ws.cell(row=last_row + i, column=10, value=order_data["order_number"])  # 新PO
        ws.cell(row=last_row + i, column=11, value=order_data.get("sales_order_number", ""))  # SO
        ws.cell(row=last_row + i, column=12, value=item["price"])  # 单价
        ws.cell(row=last_row + i, column=13, value="")  # 是否已出（留空，后续手动填写）
        ws.cell(row=last_row + i, column=14, value="")  # 对比（留空）
        ws.cell(row=last_row + i, column=15, value="")  # 标记（留空）
    
    wb.save(stats_file)
    wb.close()
    
    print(f"统计表已更新: {stats_file}")
    return True

# ============== Oracle自动化操作模块 ==============

def automate_oracle_import(template_file, coords=None):
    """
    使用pyautogui自动化Oracle导入操作
    注意：这需要Oracle客户端已经打开
    
    参数：
        template_file: 模板文件路径
        coords: 界面坐标字典，从config.py加载
    """
    try:
        import pyautogui
        import pyperclip
    except ImportError:
        print("请安装pyautogui和pyperclip: pip install pyautogui pyperclip")
        return None
    
    # 加载坐标配置
    try:
        from config import ORACLE_COORDS
        if coords is None:
            coords = ORACLE_COORDS
    except ImportError:
        print("警告: 未找到坐标配置，使用默认值")
        coords = {
            "import_button": (100, 100),
            "file_input": (150, 150),
            "confirm_button": (200, 200),
            "export_button": (250, 250),
        }
    
    print("开始Oracle自动化导入...")
    print("请确保Oracle客户端已打开并显示在屏幕上")
    print("等待10秒，请切换到Oracle窗口...")
    time.sleep(10)
    
    sales_order_number = None
    
    try:
        # 1. 激活Oracle窗口（按Alt+Tab切换）
        pyautogui.hotkey('alt', 'tab')
        time.sleep(1)
        
        # 2. 点击导入按钮
        print(f"点击导入按钮: {coords['import_button']}")
        pyautogui.click(coords['import_button'][0], coords['import_button'][1])
        time.sleep(2)
        
        # 3. 输入模板文件路径
        print(f"输入文件路径...")
        pyautogui.click(coords['file_input'][0], coords['file_input'][1])
        time.sleep(0.5)
        pyperclip.copy(str(template_file))
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1)
        
        # 4. 点击确认导入
        print(f"点击确认按钮: {coords['confirm_button']}")
        pyautogui.click(coords['confirm_button'][0], coords['confirm_button'][1])
        time.sleep(15)  # 等待导入处理
        
        # 5. 导出销售订单号
        print(f"点击导出按钮: {coords['export_button']}")
        pyautogui.click(coords['export_button'][0], coords['export_button'][1])
        time.sleep(3)
        
        # 6. 复制销售订单号（假设在某位置显示）
        # 这里需要根据实际情况调整
        # 尝试选中并复制
        pyautogui.click(coords.get('sales_order_field', (300, 300))[0],
                       coords.get('sales_order_field', (300, 300))[1])
        time.sleep(0.5)
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.5)
        sales_order_number = pyperclip.paste()
        
        print(f"获取销售订单号: {sales_order_number}")
        
    except Exception as e:
        print(f"Oracle自动化出错: {e}")
        return None
    
    print("Oracle自动化操作完成")
    return sales_order_number

# ============== 邮件发送模块 ==============

def send_result_email(result_data):
    """发送结果报告邮件"""
    import smtplib
    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart
    
    print("正在发送结果邮件...")
    
    msg = MIMEMultipart()
    msg['From'] = CONFIG["email"]["account"]
    msg['To'] = CONFIG["report_email"]
    msg['Subject'] = f"采购订单自动化报告 - {date.today().strftime('%Y-%m-%d')}"
    
    body = f"""
采购订单自动化执行报告
执行时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
执行日期: {date.today().strftime('%Y-%m-%d')}

处理结果:
- 处理邮件数: {result_data['emails_processed']}
- 生成订单数: {result_data['orders_created']}
- 订单详情:
{result_data['order_details']}

状态: {result_data['status']}

如有问题请及时处理。
"""
    
    msg.attach(MIMEText(body, 'plain', 'utf-8'))
    
    try:
        # SMTP发送（需要根据邮件服务商配置）
        # 这里使用IMAP服务商的SMTP，通常在同一主机
        smtp_server = CONFIG["email"]["imap_server"].replace("imap", "smtp")
        
        with smtplib.SMTP(smtp_server, 587) as server:
            server.starttls()
            server.login(CONFIG["email"]["account"], CONFIG["email"]["password"])
            server.send_message(msg)
        
        print("结果邮件已发送")
        return True
    except Exception as e:
        print(f"发送邮件失败: {e}")
        return False

# ============== 主流程 ==============

def main():
    """主执行流程"""
    print("="*60)
    print(f"采购订单自动化脚本启动 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*60)
    
    result_data = {
        "emails_processed": 0,
        "orders_created": 0,
        "order_details": "",
        "status": "成功",
        "sales_order_numbers": []
    }
    
    try:
        # 1. 读取今日邮件
        print("\n[1] 读取今日邮件...")
        emails = get_today_emails_from_source()
        result_data["emails_processed"] = len(emails)
        
        if not emails:
            print("今日无新邮件，退出")
            result_data["status"] = "无新邮件"
            send_result_email(result_data)
            return
        
        # 2. 加载价格表
        print("\n[2] 加载价格表...")
        prices = load_price_table()
        
        # 3. 解析邮件并创建订单
        print("\n[3] 解析邮件内容...")
        
        all_orders = []
        for email_data in emails:
            order_info = parse_order_info_from_email(email_data["body"], email_data["subject"])
            
            if not order_info["items"]:
                print(f"邮件 '{email_data['subject']}' 未解析出订单信息，跳过")
                continue
            
            if not order_info["entity"]:
                print("警告: 未检测到主体，默认使用SZK")
                order_info["entity"] = "SZK"
            
            # 匹配价格
            for item in order_info["items"]:
                item["price"] = match_price(item["model"], prices)
            
            all_orders.append(order_info)
        
        if not all_orders:
            print("无有效订单信息，退出")
            result_data["status"] = "无有效订单"
            send_result_email(result_data)
            return
        
        # 4. 生成订单号并填充模板
        print("\n[4] 生成采购订单...")
        
        for order_info in all_orders:
            entity = order_info["entity"]
            
            # 获取现有订单号避免重复
            existing_numbers = get_existing_order_numbers(
                STATISTICS_DIR / "domestic_statistics.xlsx"
            ) + get_existing_order_numbers(
                STATISTICS_DIR / "international_statistics.xlsx"
            )
            
            order_number = generate_order_number(entity, existing_numbers)
            order_info["order_number"] = order_number
            
            print(f"订单号: {order_number}")
            print(f"主体: {ENTITY_CONFIG[entity]['name']}")
            print(f"币种: {ENTITY_CONFIG[entity]['currency']}")
            print(f"型号数量:")
            for item in order_info["items"]:
                print(f"  - {item['model']}: {item['quantity']}pcs @ {item['price']}")
            
            # 填充模板
            output_file = OUTPUT_DIR / f"{order_number}.xlsm"
            fill_webadi_template(order_info, output_file)
            
            # 5. Oracle导入（需要Oracle客户端已打开）
            print("\n[5] Oracle导入...")
            sales_order_number = automate_oracle_import(output_file)
            if sales_order_number:
                order_info["sales_order_number"] = sales_order_number
                print(f"销售订单号: {sales_order_number}")
            
            # 6. 更新统计表
            print("\n[6] 更新统计表...")
            update_statistics_table(order_info, entity)
            
            result_data["orders_created"] += 1
            result_data["order_details"] += f"\n订单号: {order_number}\n"
            for item in order_info["items"]:
                result_data["order_details"] += f"  {item['model']}: {item['quantity']}pcs\n"
        
        # 7. 发送结果邮件
        print("\n[7] 发送结果邮件...")
        send_result_email(result_data)
        
        print("\n" + "="*60)
        print("采购订单自动化完成")
        print("="*60)
        
    except Exception as e:
        print(f"\n执行出错: {e}")
        import traceback
        traceback.print_exc()
        result_data["status"] = f"失败: {str(e)}"
        send_result_email(result_data)

if __name__ == "__main__":
    main()
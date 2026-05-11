# 采购订单自动化配置文件
# 请根据实际情况修改

# 邮箱配置
EMAIL_ACCOUNT = "yuchuan.he@casue.com"
EMAIL_PASSWORD = "-DxpOD5kkN)(RuPgAK-p"
IMAP_SERVER = "imap.appia.vip"
SOURCE_EMAIL = "na.yang_w@casue.com"  # 监控的发件人

# ERP配置
ERP_USERNAME = "607693"
ERP_PASSWORD = "hyc010815"

# 报告发送目标
REPORT_EMAIL = "yuchuan.he@casue.com"

# 主体映射（可根据需要添加）
ENTITY_MAP = {
    "SZK": {"name": "世纪云芯", "currency": "CNY"},
    "ICK": {"name": "智能云芯", "currency": "CNY"},
    "HSJ": {"name": "海南世纪", "currency": "CNY"},
    "DPT": {"name": "Bitmain Development PTE. LTD.", "currency": "USD"},
    "BJK": {"name": "Bitmain Beijing", "currency": "CNY"},
}

# 物料编码映射（根据实际编码规则调整）
# 格式: "型号": "编码"
MATERIAL_CODE_MAP = {
    "BM1362AA": "Y09BM1362A00",
    "BM1362AC": "Y09BM1362C00",
    "BM1362AK": "Y09BM1362310",
    "BM1366": "Y09BM1366820",
    "BM1398": "Y09BM1398240",
    "BM1489": "Y09BM1489010",
    "BM1746": "Y09BM1746010",
    "BM1684": "Y09BM1684440",
    # 更多编码请根据实际添加...
}

# Oracle界面坐标配置（需根据你的屏幕调整）
ORACLE_COORDS = {
    "import_button": (100, 100),  # 导入按钮位置
    "file_input": (150, 150),     # 文件路径输入框
    "confirm_button": (200, 200), # 确认按钮
    "export_button": (250, 250),  # 导出按钮
}
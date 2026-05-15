import pandas as pd
import glob
import os
import warnings
warnings.filterwarnings('ignore')

# ---------------------- 全局配置（可自行修改） ----------------------
# 定义芯片型号对应的gross_die值
gross_die = {"BM1370": 2614, "BM1366": 4255, "BM1368": 3186,"BM1340":1027,"BM1373":798,"BM1493":907}

# ====================== 【核心新增】外协名称标准化映射 ======================
osat_alias_map = {
    "ASEKH": ["ASE-KH", "ASEKH", "ASESH", "ATXSH"],
    "JSCC": ["JSCC"],
    "NJVT": ["VTEST", "NJ-VT", "VT", "NJVT"],
    "SPILSZ": ["SPILSZ"],
    "SPILTW": ["SPILTW","SPIL-CH", "SPILCH"],
    "TFME": ["TFME", "TFME/STTF","STTF"],
    "WX-HN": [
        "HN", "TT-HN", "WX-HN",
        "Jiangsu Haina Electronics Technology Co., Ltd."
    ],
    "WINSTEK": ["Winstek", "Winstek DPS"],
    "XJ": ["XJ", "Zhenjiang Silicon Exceed TestingTechnology Co."]
}

# 补齐后的复盘时间字典（标准名称为Key）
review_time_str = {
    "WINSTEK": "2026/1/16",
    "SPILTW": "2026/1/13",
    "ASEKH": "2026/1/13",
    "ASECL": "2026/1/12",
    "TFME": "2026/1/12",
    "SPILSZ": "2026/1/12",
    "JSCC": "2026/1/15",
    "XJ": "2026/1/19",
    "WX-HN": "2026/2/2",
    "NJVT": "2026/3/4"

}

# 转换为datetime用于时间比较
review_time = {
    k: pd.to_datetime(v, format="%Y/%m/%d") for k, v in review_time_str.items()
}

# 【全局可配置】数据筛选截止时间
CUTOFF_TIME = pd.to_datetime("2026-01-29 00:00:00")
# 基础路径
base_path = "E:\\供应链\\芯片\\"
# 时间统一格式
TIME_FORMAT = "%Y-%m-%d %H:%M:%S"

# ====================== 【核心新增】外协名称标准化函数 ======================
def standardize_osat_name(osat_name):
    """
    外协名称标准化：自动匹配别名、处理大小写/符号/空格
    """
    if pd.isna(osat_name) or not isinstance(osat_name, str):
        return osat_name
    
    cleaned_name = osat_name.strip().upper().replace("-", "").replace(" ", "")
    
    for std_name, alias_list in osat_alias_map.items():
        for alias in alias_list:
            cleaned_alias = alias.strip().upper().replace("-", "").replace(" ", "")
            if cleaned_alias == cleaned_name:
                return std_name
    return osat_name

# ====================== 通用时间处理函数 ======================
def convert_time_col(df: pd.DataFrame, col_name: str) -> pd.DataFrame:
    """统一转换时间列为datetime，剔除无效时间行"""
    if col_name not in df.columns:
        raise ValueError(f"表格中未找到时间字段「{col_name}」，请检查列名！")
    df[col_name] = pd.to_datetime(df[col_name], errors="coerce")
    df = df[pd.notna(df[col_name])].reset_index(drop=True)
    print(f"✅ 时间字段「{col_name}」处理完成，剩余 {len(df)} 行有效数据")
    return df

# ====================== OSAT时间过滤函数（已适配标准化名称） ======================
def filter_osat_time(df: pd.DataFrame, osat_col: str, time_col: str) -> pd.DataFrame:
    """
    先标准化OSAT名称，再按复盘时间过滤
    """
    # 关键：先对OSAT列做标准化
    df[osat_col] = df[osat_col].apply(standardize_osat_name)
    
    df_known = df[df[osat_col].isin(review_time.keys())].copy()
    df_unknown = df[~df[osat_col].isin(review_time.keys())].copy()

    if not df_known.empty:
        df_known["review_time"] = df_known[osat_col].map(review_time)
        df_known = df_known[df_known[time_col] < df_known["review_time"]]
        df_known = df_known.drop(columns=["review_time"])
    
    df_filtered = pd.concat([df_known, df_unknown], ignore_index=True)
    print(f"✅ OSAT时间过滤完成：有效保留 {len(df_filtered)} 行")
    return df_filtered

# ====================== 主处理函数 ======================
def process_chip_model(chip_model: str, die_value: int):
    chip_dir = os.path.join(base_path, chip_model)
    if not os.path.exists(chip_dir):
        print(f"⚠️ 芯片型号 {chip_model} 目录不存在：{chip_dir}\n")
        return

    # 查找文件
    file_patterns = {
        "wafer": os.path.join(chip_dir, "IcopWaferShipping*.xlsx"),
        "bumpcp": os.path.join(chip_dir, f"{chip_model}_bumping_cp_shipping*.xlsx"),
        "assembly": os.path.join(chip_dir, f"{chip_model}_assembly_shipping*.xlsx"),
        "test": os.path.join(chip_dir, f"{chip_model}_FACTORY_WO*.xlsx")
    }

    file_paths = {}
    for key, pattern in file_patterns.items():
        matched = glob.glob(pattern)
        if matched:
            file_paths[key] = matched[0]
            print(f"✅ {chip_model} - {key} 找到：{os.path.basename(file_paths[key])}")
        else:
            print(f"❌ {chip_model} - {key} 未找到：{pattern}\n")
            return

    # 读取并处理各表格
    try:
        # 表1 Wafer
        df1 = pd.read_excel(file_paths["wafer"])
        print(f"\n📄 {chip_model} - wafer 原始行数：{len(df1)}")
        df1 = convert_time_col(df1, "创建时间")
        df1 = df1[df1["创建时间"] < CUTOFF_TIME].reset_index(drop=True)
        df1 = filter_osat_time(df1, "Ship Code", "创建时间")
        # 表2 BumpCP：创建时间 + ship_date 标准化 + 过滤
        df2 = pd.read_excel(file_paths["bumpcp"])
        print(f"\n📄 {chip_model} - bumpcp 原始行数：{len(df2)}")
        df2 = convert_time_col(df2, "shippend_date")
        df2 = df2[df2["shippend_date"] < CUTOFF_TIME].reset_index(drop=True)
        df2 = filter_osat_time(df2, "osat", "shippend_date")
        df2 = filter_osat_time(df2, "ship_code", "shippend_date")

        # 表3 Assembly：shippend_date 标准化 + 过滤
        df3 = pd.read_excel(file_paths["assembly"])
        print(f"\n📄 {chip_model} - assembly 原始行数：{len(df3)}")
        df3 = convert_time_col(df3, "ship_date")
        df3 = df3[df3["ship_date"] < CUTOFF_TIME].reset_index(drop=True)
        df3 = filter_osat_time(df3, "osat", "ship_date")
        df3 = filter_osat_time(df3, "ship_to", "ship_date")

        # 表4 Test：merge_date 标准化 + 过滤
        df4 = pd.read_excel(file_paths["test"])
        print(f"\n📄 {chip_model} - test 原始行数：{len(df4)}")
        df4 = convert_time_col(df4, "merge_date")
        df4 = df4[df4["merge_date"] < CUTOFF_TIME].reset_index(drop=True)
        df4 = filter_osat_time(df4, "test_supplier", "merge_date")

    except Exception as e:
        print(f"\n❌ {chip_model} 表格处理失败：{str(e)}\n")
        return

    # 初始化统计字典
    hold_dict = {}
    out_only_dict = {}

    # ---------------------- 处理表1 Wafer ----------------------
    for _, row in df1.iterrows():
        wafer_id = row["Wafer Lot"]
        wafer_pn = row["Wafer Pn"][:6]
        send_factory = "台积电"
        receive_factory = standardize_osat_name(row["Ship Code"])  # 接收方也标准化
        qty = row["Wafer Qty"] * die_value
        stage = "wafer"

        key_send = (wafer_id, wafer_pn, send_factory, stage)
        hold_dict[key_send] = hold_dict.get(key_send, 0) - qty
        out_only_dict[key_send] = out_only_dict.get(key_send, 0) + qty

        key_receive = (wafer_id, wafer_pn, receive_factory, stage)
        hold_dict[key_receive] = hold_dict.get(key_receive, 0) + qty

    # ---------------------- 处理表2 BumpCP ----------------------
    for _, row in df2.iterrows():
        wafer_id = row["wafer_lot_id"]
        wafer_pn = row["wafer_pn"][:6]
        send_factory = row["osat"]  # 已在filter中标准化
        receive_factory = standardize_osat_name(row["ship_code"])
        qty = row["chip_qty"]
        stage = row["factory"] if pd.notna(row["factory"]) else "bumpcp"

        key_send = (wafer_id, wafer_pn, send_factory, stage)
        hold_dict[key_send] = hold_dict.get(key_send, 0) - qty
        out_only_dict[key_send] = out_only_dict.get(key_send, 0) + qty

        key_receive = (wafer_id, wafer_pn, receive_factory, stage)
        hold_dict[key_receive] = hold_dict.get(key_receive, 0) + qty

    # ---------------------- 处理表3 Assembly ----------------------
    for _, row in df3.iterrows():
        wafer_id = row["wafer_lot_id"]
        wafer_pn = row["device_pn"][:6]
        send_factory = row["osat"]  # 已标准化
        receive_factory = standardize_osat_name(row["ship_to"])
        qty = row["ship_qty"]
        stage = row["stage"] if pd.notna(row["stage"]) else "assembly/test"

        key_send = (wafer_id, wafer_pn, send_factory, stage)
        hold_dict[key_send] = hold_dict.get(key_send, 0) - qty
        out_only_dict[key_send] = out_only_dict.get(key_send, 0) + qty

        key_receive = (wafer_id, wafer_pn, receive_factory, stage)
        hold_dict[key_receive] = hold_dict.get(key_receive, 0) + qty

    # ---------------------- 处理表4 Test ----------------------
    for _, row in df4.iterrows():
        test_supplier = row["test_supplier"]  # 已标准化
        wafer_lot = row["wafer_lot"]
        mix_marking = standardize_osat_name(row["mix_marking"])
        qty = row["qty"]

        if pd.notna(row["device_name"]):
            wafer_pn = row["device_name"][:6]
        else:
            wafer_pn = row["product_name"][:6] if pd.notna(row["product_name"]) else chip_model
        stage = 1

        key_reduce = (wafer_lot, wafer_pn, test_supplier, stage)
        hold_dict[key_reduce] = hold_dict.get(key_reduce, 0) - qty
        out_only_dict[key_reduce] = out_only_dict.get(key_reduce, 0) + qty

        key_increase = (mix_marking, wafer_pn, test_supplier, stage)
        hold_dict[key_increase] = hold_dict.get(key_increase, 0) + qty

    # 生成汇总表
    summary_list = []
    for (wafer_id, wafer_pn, osat_factory, stage), hold_qty in hold_dict.items():
        summary_list.append({
            "wafer_id": wafer_id,
            "wafer_pn": wafer_pn,
            "osat_factory": osat_factory,  # 输出为标准名称
            "stage": stage,
            "hold_qty": hold_qty,
            "外协对应阶段只出的数目(不算入)": out_only_dict.get((wafer_id, wafer_pn, osat_factory, stage), 0)
        })

    df_summary = pd.DataFrame(summary_list)
    df_summary = df_summary.sort_values(by=["wafer_pn", "stage", "osat_factory"], ignore_index=True)

    output_path = os.path.join(chip_dir, f"{chip_model}_芯片阶段数量汇总表_0112.xlsx")
    df_summary.to_excel(output_path, index=False, engine="openpyxl")

    print(f"\n🎉 {chip_model} 汇总表生成完成：{output_path}")
    print(f"📊 最终有效汇总行数：{len(df_summary)}\n{'-'*60}")

# ---------------------- 主程序入口 ----------------------
if __name__ == "__main__":
    print("="*70)
    print("开始执行芯片供应链数据处理（已开启外协名称标准化）")
    print(f"全局截止时间：{CUTOFF_TIME.strftime(TIME_FORMAT)}")
    print("="*70, "\n")

    for model, die in gross_die.items():
        process_chip_model(model, die)

    print("="*70)
    print("所有芯片型号处理完毕！")
    print("="*70)
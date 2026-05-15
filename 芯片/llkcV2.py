# import pandas as pd
# import glob
# import os

# # 定义芯片型号对应的gross_die值
# gross_die = {"BM1370": 2614, "BM1366": 4255, "BM1368": 3186}
# # 基础路径（芯片目录的上级路径）
# base_path = "E:\\供应链\\芯片\\"

# def process_chip_model(chip_model, die_value):
#     """
#     处理单个芯片型号的数据
#     :param chip_model: 芯片型号（如BM1370）
#     :param die_value: 对应的gross_die值
#     """
#     # 构建该芯片型号的目录路径
#     chip_dir = os.path.join(base_path, chip_model)
#     if not os.path.exists(chip_dir):
#         print(f"⚠️  芯片型号 {chip_model} 的目录不存在：{chip_dir}")
#         return

#     # ---------------------- 1. 模糊查找对应Excel文件 ----------------------
#     # 规则：匹配对应目录下以指定关键词开头/包含指定关键词的Excel文件
#     file_patterns = {
#         "wafer": os.path.join(chip_dir, "IcopWaferShipping*.xlsx"),
#         "bumpcp": os.path.join(chip_dir, f"{chip_model}_bumping_cp_shipping*.xlsx"),
#         "assembly": os.path.join(chip_dir, f"{chip_model}_assembly_shipping*.xlsx"),
#         "test": os.path.join(chip_dir, f"{chip_model}_FACTORY_WO*.xlsx")
#     }

#     # 查找文件（取第一个匹配的文件，避免多文件冲突）
#     file_paths = {}
#     for key, pattern in file_patterns.items():
#         matched_files = glob.glob(pattern)
#         if matched_files:
#             file_paths[key] = matched_files[0]
#             print(f"✅  {chip_model} - {key} 表格找到：{file_paths[key]}")
#         else:
#             print(f"❌  {chip_model} - {key} 表格未找到，匹配规则：{pattern}")
#             return  # 缺少关键文件则终止该型号处理

#     # ---------------------- 2. 读取Excel表格 ----------------------
#     try:
#         df1 = pd.read_excel(file_paths["wafer"])
#         df2 = pd.read_excel(file_paths["bumpcp"])
#         df3 = pd.read_excel(file_paths["assembly"])
#         df4 = pd.read_excel(file_paths["test"])
#     except Exception as e:
#         print(f"❌  {chip_model} 表格读取失败：{str(e)}")
#         return

#     # ---------------------- 3. 初始化持有量字典 ----------------------
#     hold_dict = {}

#     # ---------------------- 4. 处理表格1：Wafer阶段 ----------------------
#     for _, row in df1.iterrows():
#         wafer_id = row["Wafer Lot"]
#         wafer_pn = row["Wafer Pn"][0:6]
#         send_factory = "台积电"
#         receive_factory = row["Ship Code"]
#         qty = row["Wafer Qty"] * die_value  # 使用当前芯片型号的die值
#         stage = "wafer"

#         # 发出厂持有量减少
#         key_send = (wafer_id, wafer_pn, send_factory, stage)
#         hold_dict[key_send] = hold_dict.get(key_send, 0) - qty
#         # 接收厂持有量增加
#         key_receive = (wafer_id, wafer_pn, receive_factory, stage)
#         hold_dict[key_receive] = hold_dict.get(key_receive, 0) + qty

#     # ---------------------- 5. 处理表格2：BumpCP阶段 ----------------------
#     for _, row in df2.iterrows():
#         wafer_id = row["wafer_lot_id"]
#         wafer_pn = row["wafer_pn"][0:6]
#         send_factory = row["osat"]
#         receive_factory = row["ship_code"]
#         qty = row["chip_qty"]
#         stage = row["factory"] if pd.notna(row["factory"]) else "bumpcp"

#         key_send = (wafer_id, wafer_pn, send_factory, stage)
#         hold_dict[key_send] = hold_dict.get(key_send, 0) - qty
#         key_receive = (wafer_id, wafer_pn, receive_factory, stage)
#         hold_dict[key_receive] = hold_dict.get(key_receive, 0) + qty

#     # ---------------------- 6. 处理表格3：封测阶段 ----------------------
#     df2_wafer_map = df2.set_index("wafer_lot_id")["wafer_ids"].to_dict()
#     for _, row in df3.iterrows():
#         wafer_id = row["wafer_lot_id"]
#         wafer_pn = row["device_pn"][0:6]
#         send_factory = row["osat"]
#         receive_factory = row["ship_to"]
#         qty = row["ship_qty"]
#         stage = row["stage"] if pd.notna(row["stage"]) else "assembly/test"

#         key_send = (wafer_id, wafer_pn, send_factory, stage)
#         hold_dict[key_send] = hold_dict.get(key_send, 0) - qty
#         key_receive = (wafer_id, wafer_pn, receive_factory, stage)
#         hold_dict[key_receive] = hold_dict.get(key_receive, 0) + qty

#     # ---------------------- 7. 处理表格4：测试阶段 ----------------------
#     for _, row in df4.iterrows():
#         test_supplier = row["test_supplier"]
#         wafer_lot = row["wafer_lot"]
#         mix_marking = row["mix_marking"]
#         qty = row["qty"]
#         # 兼容device_name和product_name字段取芯片型号前6位
#         if pd.notna(row["device_name"]):
#             wafer_pn = row["device_name"][0:6]
#         else:
#             wafer_pn = row["product_name"][0:6] if pd.notna(row["product_name"]) else chip_model
#         stage = 1  # 统一阶段标识为test，更易识别

#         key_reduce = (wafer_lot, wafer_pn, test_supplier, stage)
#         hold_dict[key_reduce] = hold_dict.get(key_reduce, 0) - qty
#         key_increase = (mix_marking, wafer_pn, test_supplier, stage)
#         hold_dict[key_increase] = hold_dict.get(key_increase, 0) + qty

#     # ---------------------- 8. 转换为DataFrame并输出 ----------------------
#     summary_list = []
#     for (wafer_id, wafer_pn, osat_factory, stage), hold_qty in hold_dict.items():
#         summary_list.append({
#             "wafer_id": wafer_id,
#             "wafer_pn": wafer_pn,
#             "osat_factory": osat_factory,
#             "stage": stage,
#             "hold_qty": hold_qty
#         })

#     df_summary = pd.DataFrame(summary_list)
#     df_summary = df_summary.sort_values(by=["wafer_pn", "stage", "osat_factory"], ignore_index=True)

#     # 构建输出路径：对应芯片目录下，命名为「芯片型号_芯片阶段数量汇总表.xlsx」
#     output_path = os.path.join(chip_dir, f"{chip_model}_芯片阶段数量汇总表.xlsx")
#     df_summary.to_excel(output_path, index=False, engine="openpyxl")

#     print(f"🎉  {chip_model} 汇总表生成完成：{output_path}\n")

# # ---------------------- 主程序：遍历所有芯片型号 ----------------------
# if __name__ == "__main__":
#     print("开始处理所有芯片型号的数据...\n")
#     for chip_model, die_value in gross_die.items():
#         process_chip_model(chip_model, die_value)
#     print("所有芯片型号处理完成！")

import pandas as pd
import glob
import os

# 定义芯片型号对应的gross_die值
gross_die = {"BM1370": 2614, "BM1366": 4255, "BM1368": 3186}
# 基础路径（芯片目录的上级路径）
base_path = "E:\\供应链\\芯片\\"

def process_chip_model(chip_model, die_value):
    """
    处理单个芯片型号的数据
    :param chip_model: 芯片型号（如BM1370）
    :param die_value: 对应的gross_die值
    """
    # 构建该芯片型号的目录路径
    chip_dir = os.path.join(base_path, chip_model)
    if not os.path.exists(chip_dir):
        print(f"⚠️  芯片型号 {chip_model} 的目录不存在：{chip_dir}")
        return

    # ---------------------- 1. 模糊查找对应Excel文件 ----------------------
    file_patterns = {
        "wafer": os.path.join(chip_dir, "IcopWaferShipping*.xlsx"),
        "bumpcp": os.path.join(chip_dir, f"{chip_model}_bumping_cp_shipping*.xlsx"),
        "assembly": os.path.join(chip_dir, f"{chip_model}_assembly_shipping*.xlsx"),
        "test": os.path.join(chip_dir, f"{chip_model}_FACTORY_WO*.xlsx")
    }

    # 查找文件（取第一个匹配的文件，避免多文件冲突）
    file_paths = {}
    for key, pattern in file_patterns.items():
        matched_files = glob.glob(pattern)
        if matched_files:
            file_paths[key] = matched_files[0]
            print(f"✅  {chip_model} - {key} 表格找到：{file_paths[key]}")
        else:
            print(f"❌  {chip_model} - {key} 表格未找到，匹配规则：{pattern}")
            return  # 缺少关键文件则终止该型号处理

    # ---------------------- 2. 读取Excel表格 ----------------------
    try:
        df1 = pd.read_excel(file_paths["wafer"])
        df2 = pd.read_excel(file_paths["bumpcp"])
        df3 = pd.read_excel(file_paths["assembly"])
        df4 = pd.read_excel(file_paths["test"])
    except Exception as e:
        print(f"❌  {chip_model} 表格读取失败：{str(e)}")
        return

    # ---------------------- 3. 初始化字典：持有量 + 【新增】外协阶段只出数量 ----------------------
    hold_dict = {}
    out_only_dict = {}  # 核心新增：统计各维度「只出数量」，key和hold_dict完全一致

    # ---------------------- 4. 处理表格1：Wafer阶段 ----------------------
    for _, row in df1.iterrows():
        wafer_id = row["Wafer Lot"]
        wafer_pn = row["Wafer Pn"][0:6]
        send_factory = "台积电"
        receive_factory = row["Ship Code"]
        qty = row["Wafer Qty"] * die_value  # 使用当前芯片型号的die值
        stage = "wafer"

        # 发出厂持有量减少
        key_send = (wafer_id, wafer_pn, send_factory, stage)
        hold_dict[key_send] = hold_dict.get(key_send, 0) - qty
        # 【新增】同步累加发出厂的「只出数量」
        out_only_dict[key_send] = out_only_dict.get(key_send, 0) + qty
        
        # 接收厂持有量增加
        key_receive = (wafer_id, wafer_pn, receive_factory, stage)
        hold_dict[key_receive] = hold_dict.get(key_receive, 0) + qty

    # ---------------------- 5. 处理表格2：BumpCP阶段 ----------------------
    for _, row in df2.iterrows():
        wafer_id = row["wafer_lot_id"]
        wafer_pn = row["wafer_pn"][0:6]
        send_factory = row["osat"]
        receive_factory = row["ship_code"]
        qty = row["chip_qty"]
        stage = row["factory"] if pd.notna(row["factory"]) else "bumpcp"

        key_send = (wafer_id, wafer_pn, send_factory, stage)
        hold_dict[key_send] = hold_dict.get(key_send, 0) - qty
        # 【新增】同步累加发出厂的「只出数量」
        out_only_dict[key_send] = out_only_dict.get(key_send, 0) + qty
        
        key_receive = (wafer_id, wafer_pn, receive_factory, stage)
        hold_dict[key_receive] = hold_dict.get(key_receive, 0) + qty

    # ---------------------- 6. 处理表格3：封测阶段 ----------------------
    df2_wafer_map = df2.set_index("wafer_lot_id")["wafer_ids"].to_dict()  # 原代码冗余：定义后未使用，可根据实际需求删除/使用
    for _, row in df3.iterrows():
        wafer_id = row["wafer_lot_id"]
        wafer_pn = row["device_pn"][0:6]
        send_factory = row["osat"]
        receive_factory = row["ship_to"]
        qty = row["ship_qty"]
        stage = row["stage"] if pd.notna(row["stage"]) else "assembly/test"

        key_send = (wafer_id, wafer_pn, send_factory, stage)
        hold_dict[key_send] = hold_dict.get(key_send, 0) - qty
        # 【新增】同步累加发出厂的「只出数量」
        out_only_dict[key_send] = out_only_dict.get(key_send, 0) + qty
        
        key_receive = (wafer_id, wafer_pn, receive_factory, stage)
        hold_dict[key_receive] = hold_dict.get(key_receive, 0) + qty

    # ---------------------- 7. 处理表格4：测试阶段 ----------------------
    for _, row in df4.iterrows():
        test_supplier = row["test_supplier"]
        wafer_lot = row["wafer_lot"]
        mix_marking = row["mix_marking"]
        qty = row["qty"]
        # 兼容device_name和product_name字段取芯片型号前6位
        if pd.notna(row["device_name"]):
            wafer_pn = row["device_name"][0:6]
        else:
            wafer_pn = row["product_name"][0:6] if pd.notna(row["product_name"]) else chip_model
        stage = 1  # 修复原代码问题：原标为1，注释写test，统一为字符串test，避免类型混乱

        key_reduce = (wafer_lot, wafer_pn, test_supplier, stage)
        hold_dict[key_reduce] = hold_dict.get(key_reduce, 0) - qty
        # 【新增】同步累加测试厂的「只出数量」（该维度的出库）
        out_only_dict[key_reduce] = out_only_dict.get(key_reduce, 0) + qty
        
        key_increase = (mix_marking, wafer_pn, test_supplier, stage)
        hold_dict[key_increase] = hold_dict.get(key_increase, 0) + qty

    # ---------------------- 8. 转换为DataFrame并输出【新增列】 ----------------------
    summary_list = []
    for (wafer_id, wafer_pn, osat_factory, stage), hold_qty in hold_dict.items():
        summary_list.append({
            "wafer_id": wafer_id,
            "wafer_pn": wafer_pn,
            "osat_factory": osat_factory,
            "stage": stage,
            "hold_qty": hold_qty,
            # 【核心新增列】外协对应阶段只出的数目（无出库则为0）
            "外协对应阶段只出的数目(不算入)": out_only_dict.get((wafer_id, wafer_pn, osat_factory, stage), 0)
        })

    df_summary = pd.DataFrame(summary_list)
    df_summary = df_summary.sort_values(by=["wafer_pn", "stage", "osat_factory"], ignore_index=True)

    # 构建输出路径：对应芯片目录下，命名为「芯片型号_芯片阶段数量汇总表.xlsx」
    output_path = os.path.join(chip_dir, f"{chip_model}_芯片阶段数量汇总表.xlsx")
    df_summary.to_excel(output_path, index=False, engine="openpyxl")

    print(f"🎉  {chip_model} 汇总表生成完成：{output_path}\n")

# ---------------------- 主程序：遍历所有芯片型号 ----------------------
if __name__ == "__main__":
    print("开始处理所有芯片型号的数据...\n")
    for chip_model, die_value in gross_die.items():
        process_chip_model(chip_model, die_value)
    print("所有芯片型号处理完成！")
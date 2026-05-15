


# import pandas as pd
# gross_die={"BM1370":2614}
# # ---------------------- 1. 读取三个Excel表格 ----------------------
# # 请修改为你的表格实际路径
# path1 = "E:\\供应链\\芯片\\BM1370\\IcopWaferShipping_20260123154038.xlsx"  # Wafer阶段表格
# path2 = "E:\\供应链\\芯片\\BM1370\\BM1370_bumping_cp_shipping_20260112113757.xlsx"  # BumpCP阶段表格
# path3 = "E:\\供应链\\芯片\\BM1370\\BM1370_assembly_shipping_20260112113757.xlsx"  # 封测阶段表格




# df1 = pd.read_excel(path1)
# df2 = pd.read_excel(path2)
# df3 = pd.read_excel(path3)

# # ---------------------- 2. 初始化持有量字典（键：四维组合，值：累计数量） ----------------------
# # 键格式：(wafer_id, wafer_pn, osat_factory, stage)
# hold_dict = {}

# # ---------------------- 3. 处理表格1：Wafer阶段（台积电→Ship Code） ----------------------
# for _, row in df1.iterrows():
#     wafer_id = row["Wafer Lot"]
#     wafer_pn = row["Wafer Pn"][0:6]
#     send_factory = "台积电"
#     receive_factory = row["Ship Code"]
#     qty = row["Wafer Qty"]*gross_die[wafer_pn]
#     stage = "wafer"

#     # 1. 发出厂（台积电）：该阶段持有量减少
#     key_send = (wafer_id, wafer_pn, send_factory, stage)
#     hold_dict[key_send] = hold_dict.get(key_send, 0) - qty

#     # 2. 接收厂（Ship Code）：该阶段持有量增加
#     key_receive = (wafer_id, wafer_pn, receive_factory, stage)
#     hold_dict[key_receive] = hold_dict.get(key_receive, 0) + qty

# # ---------------------- 4. 处理表格2：BumpCP阶段（osat→ship_code） ----------------------
# for _, row in df2.iterrows():
#     wafer_id = row["wafer_lot_id"]
#     wafer_pn = row["wafer_pn"][0:6]
#     send_factory = row["osat"]
#     receive_factory = row["ship_code"]
#     qty = row["chip_qty"]
#     stage = row["factory"] if pd.notna(row["factory"]) else "bumpcp"  # 空值填充

#     # 1. 发出厂（osat）：该阶段持有量减少
#     key_send = (wafer_id, wafer_pn, send_factory, stage)
#     hold_dict[key_send] = hold_dict.get(key_send, 0) - qty

#     # 2. 接收厂（ship_code）：该阶段持有量增加
#     key_receive = (wafer_id, wafer_pn, receive_factory, stage)
#     hold_dict[key_receive] = hold_dict.get(key_receive, 0) + qty

# # ---------------------- 5. 处理表格3：封测阶段（osat→ship_to） ----------------------
# # 先建立表格2的wafer_lot_id→wafer_id映射
# df2_wafer_map = df2.set_index("wafer_lot_id")["wafer_ids"].to_dict()

# for _, row in df3.iterrows():
#     wafer_id = row["wafer_lot_id"]
#     wafer_pn = row["device_pn"][0:6]
#     # 关联wafer_id，无匹配则为"未知"
#     send_factory = row["osat"]
#     receive_factory = row["ship_to"]
#     qty = row["ship_qty"]
#     stage = row["stage"] if pd.notna(row["stage"]) else "assembly/test"  # 空值填充

#     # 1. 发出厂（osat）：该阶段持有量减少
#     key_send = (wafer_id, wafer_pn, send_factory, stage)
#     hold_dict[key_send] = hold_dict.get(key_send, 0) - qty

#     # 2. 接收厂（ship_to）：该阶段持有量增加
#     key_receive = (wafer_id, wafer_pn, receive_factory, stage)
#     hold_dict[key_receive] = hold_dict.get(key_receive, 0) + qty

# # ---------------------- 6. 转换为DataFrame并整理输出 ----------------------
# # 拆解字典为列表
# summary_list = []
# for (wafer_id, wafer_pn, osat_factory, stage), hold_qty in hold_dict.items():
#     summary_list.append({
#         "wafer_id": wafer_id,
#         "wafer_pn": wafer_pn,
#         "osat_factory": osat_factory,
#         "stage": stage,
#         "hold_qty": hold_qty
#     })

# df_summary = pd.DataFrame(summary_list)
# # 按字段排序，方便查看
# df_summary = df_summary.sort_values(by=["wafer_pn", "stage", "osat_factory"], ignore_index=True)

# # 输出到Excel
# output_path = "E:\\供应链\\芯片\\BM1370\\芯片阶段数量汇总表.xlsx"
# df_summary.to_excel(output_path, index=False, engine="openpyxl")

# print(f"最终持有量汇总表已生成：{output_path}")
# print("字段说明：")
# print("- wafer_id: 芯片Wafer唯一标识")
# print("- wafer_pn: 芯片型号")
# print("- osat_factory: 工厂/OSAT名称")
# print("- stage: 芯片所处阶段")
# print("- hold_qty: 最终持有量（正数=持有，负数=发出未匹配接收，需核对数据）")



import pandas as pd

gross_die={"BM1370":2614,"BM1366":4255,"BM1368":3186}
# ---------------------- 1. 读取四个Excel表格 ----------------------
# 请修改为你的表格实际路径
path1 = "E:\\供应链\\芯片\\BM1370\\IcopWaferShipping_20260123154038.xlsx"  # Wafer阶段表格
path2 = "E:\\供应链\\芯片\\BM1370\\BM1370_bumping_cp_shipping_20260112113757.xlsx"  # BumpCP阶段表格
path3 = "E:\\供应链\\芯片\\BM1370\\BM1370_assembly_shipping_20260112113757.xlsx"  # 封测阶段表格
# 新增：第四个测试阶段表格路径（请补充实际路径）
path4 = "E:\\供应链\\芯片\\BM1370\\BM1370_FACTORY_WO_20260126105541.xlsx"  

df1 = pd.read_excel(path1)
df2 = pd.read_excel(path2)
df3 = pd.read_excel(path3)
# 新增：读取第四个表格
df4 = pd.read_excel(path4)

# ---------------------- 2. 初始化持有量字典（键：四维组合，值：累计数量） ----------------------
# 键格式：(wafer_id/标识, wafer_pn, osat_factory, stage)
hold_dict = {}

# ---------------------- 3. 处理表格1：Wafer阶段（台积电→Ship Code） ----------------------
for _, row in df1.iterrows():
    wafer_id = row["Wafer Lot"]
    wafer_pn = row["Wafer Pn"][0:6]
    send_factory = "台积电"
    receive_factory = row["Ship Code"]
    qty = row["Wafer Qty"]*gross_die[wafer_pn]
    stage = "wafer"

    # 1. 发出厂（台积电）：该阶段持有量减少
    key_send = (wafer_id, wafer_pn, send_factory, stage)
    hold_dict[key_send] = hold_dict.get(key_send, 0) - qty

    # 2. 接收厂（Ship Code）：该阶段持有量增加
    key_receive = (wafer_id, wafer_pn, receive_factory, stage)
    hold_dict[key_receive] = hold_dict.get(key_receive, 0) + qty

# ---------------------- 4. 处理表格2：BumpCP阶段（osat→ship_code） ----------------------
for _, row in df2.iterrows():
    wafer_id = row["wafer_lot_id"]
    wafer_pn = row["wafer_pn"][0:6]
    send_factory = row["osat"]
    receive_factory = row["ship_code"]
    qty = row["chip_qty"]
    stage = row["factory"] if pd.notna(row["factory"]) else "bumpcp"  # 空值填充

    # 1. 发出厂（osat）：该阶段持有量减少
    key_send = (wafer_id, wafer_pn, send_factory, stage)
    hold_dict[key_send] = hold_dict.get(key_send, 0) - qty

    # 2. 接收厂（ship_code）：该阶段持有量增加
    key_receive = (wafer_id, wafer_pn, receive_factory, stage)
    hold_dict[key_receive] = hold_dict.get(key_receive, 0) + qty

# ---------------------- 5. 处理表格3：封测阶段（osat→ship_to） ----------------------
# 先建立表格2的wafer_lot_id→wafer_id映射
df2_wafer_map = df2.set_index("wafer_lot_id")["wafer_ids"].to_dict()

for _, row in df3.iterrows():
    wafer_id = row["wafer_lot_id"]
    wafer_pn = row["device_pn"][0:6]
    # 关联wafer_id，无匹配则为"未知"
    send_factory = row["osat"]
    receive_factory = row["ship_to"]
    qty = row["ship_qty"]
    stage = row["stage"] if pd.notna(row["stage"]) else "assembly/test"  # 空值填充

    # 1. 发出厂（osat）：该阶段持有量减少
    key_send = (wafer_id, wafer_pn, send_factory, stage)
    hold_dict[key_send] = hold_dict.get(key_send, 0) - qty

    # 2. 接收厂（ship_to）：该阶段持有量增加
    key_receive = (wafer_id, wafer_pn, receive_factory, stage)
    hold_dict[key_receive] = hold_dict.get(key_receive, 0) + qty

# ---------------------- 6. 新增：处理表格4：测试阶段（test_supplier内部流转） ----------------------
for _, row in df4.iterrows():
    # 提取核心字段
    test_supplier = row["test_supplier"]  # 工厂名称
    wafer_lot = row["wafer_lot"]          # 减少的标识（对应原有wafer_id）
    mix_marking = row["mix_marking"]      # 增加的标识
    qty = row["qty"]                      # 变动数量
    # 芯片型号：取device_name前6位，和原有逻辑保持一致
    wafer_pn = row["device_name"][0:6] if pd.notna(row["device_name"]) else row["product_name"][0:6]
    stage = 1  # 该表格阶段固定为test，便于区分

    # 规则1：test_supplier的wafer_lot维度减少qty
    key_reduce = (wafer_lot, wafer_pn, test_supplier, stage)
    hold_dict[key_reduce] = hold_dict.get(key_reduce, 0) - qty

    # 规则2：test_supplier的mix_marking维度增加qty
    key_increase = (mix_marking, wafer_pn, test_supplier, stage)
    hold_dict[key_increase] = hold_dict.get(key_increase, 0) + qty

# ---------------------- 7. 转换为DataFrame并整理输出 ----------------------
# 拆解字典为列表
summary_list = []
for (wafer_id, wafer_pn, osat_factory, stage), hold_qty in hold_dict.items():
    summary_list.append({
        "wafer_id": wafer_id,
        "wafer_pn": wafer_pn,
        "osat_factory": osat_factory,
        "stage": stage,
        "hold_qty": hold_qty
    })

df_summary = pd.DataFrame(summary_list)
# 按字段排序，方便查看
df_summary = df_summary.sort_values(by=["wafer_pn", "stage", "osat_factory"], ignore_index=True)

# 输出到Excel
output_path = "E:\\供应链\\芯片\\BM1370\\芯片阶段数量汇总表.xlsx"
df_summary.to_excel(output_path, index=False, engine="openpyxl")

print(f"最终持有量汇总表已生成：{output_path}")
print("字段说明：")
print("- wafer_id: 芯片Wafer唯一标识（表格4中为wafer_lot/mix_marking）")
print("- wafer_pn: 芯片型号（取前6位）")
print("- osat_factory: 工厂/OSAT名称")
print("- stage: 芯片所处阶段（新增test阶段对应表格4）")
print("- hold_qty: 最终持有量（正数=持有，负数=发出未匹配接收，需核对数据）")

input_file = "E:\\供应链\\芯片\\fg_list_20251125_20251202.csv"  # 你的制表符分隔CSV文件
output_file = "E:\\供应链\\芯片\\processfg_list_20251125_20251202.xlsx"  # 输出的标准CSV
encoding = "utf-8"  # 若文件是GBK编码，改为"gbk"

import pandas as pd
# import pandas as pd

df = pd.read_csv(input_file, sep = '\t')
df.to_excel(output_file, index=False)
# 1. 读取CSV文件（指定文件路径，相对路径/绝对路径均可）
# encoding = "utf-8"  # 中文乱码时改为 "gbk"

# # 1. 读取TSV（必须指定 sep="\t"，否则数据错乱）
# df = pd.read_csv(
#     input_file,
#     encoding=encoding,
#     sep="\t",  # 核心参数：适配制表符分隔
#     na_filter=False  # 可选：保留空值原样（避免自动转换为NaN）
# )

# # 2. 保存为XLSX文件（index=False 表示不保存行索引，避免多余列）
# df.to_excel(output_file, index=False, engine="openpyxl")

# print("CSV转XLSX完成！")

# import csv
# import chardet
# import re

# # ---------------------- 步骤1：检测文件编码 ----------------------
# def get_file_encoding(file_path):
#     """检测文件编码并返回"""
#     with open(file_path, "rb") as f:
#         raw_data = f.read(1024 * 1024)
#         result = chardet.detect(raw_data)
#         return result["encoding"] or "utf-8"  # 无结果时默认utf-8

# # ---------------------- 步骤2：转换TSV到CSV ----------------------
# tsv_file = input_file  # 你的TSV文件路径
# csv_file = output_file  # 转换后的CSV文件路径
# file_encoding = get_file_encoding(tsv_file)  # 获取实际编码
# print(f"使用编码：{file_encoding} 读取文件")

# # 打开文件并转换
# with open(tsv_file, "r", encoding=file_encoding, newline="") as tsv_f, \
#      open(csv_file, "w", encoding="utf-8", newline="") as csv_f:  # CSV建议用utf-8保存
#     # 配置TSV读取器：处理多制表符/制表符+空格，同时支持引号包裹的特殊数据
#     # skipinitialspace=True：忽略分隔符后的空格；quoting=csv.QUOTE_ALL：强制用引号包裹字段
#     tsv_reader = csv.reader(
#         (re.sub(r"\t+", "\t", line.strip()) for line in tsv_f),  # 多制表符替换为单制表符
#         delimiter="\t",
#         skipinitialspace=True,  # 忽略制表符后的空格
#         quoting=csv.QUOTE_MINIMAL  # 仅对含特殊字符的字段加引号
#     )
#     # 配置CSV写入器：处理含逗号的字段，自动加引号
#     csv_writer = csv.writer(
#         csv_f,
#         delimiter=",",
#         quoting=csv.QUOTE_MINIMAL,
#         escapechar="\\"  # 转义特殊字符
#     )

#     # 逐行读取并过滤无效行（空行、纯空格行）
#     for row_num, row in enumerate(tsv_reader, 1):
#         # 过滤空行：去除全为空格的字段后若为空，则跳过
#         cleaned_row = [field.strip() for field in row]
#         if not any(cleaned_row):
#             print(f"跳过第{row_num}行：空行")
#             continue
#         # 写入CSV
#         csv_writer.writerow(cleaned_row)

# print(f"转换完成！CSV文件已保存至：{csv_file}")


#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Oracle界面坐标获取工具
运行此脚本后，将鼠标移动到Oracle界面的各个按钮位置，
按Ctrl+C记录坐标，按Ctrl+Q退出

使用方法：
1. 打开Oracle客户端
2. 运行此脚本: python get_oracle_coords.py
3. 将鼠标移动到目标位置，按Ctrl+C记录
4. 完成后按Ctrl+Q退出
5. 将记录的坐标填入config.py的ORACLE_COORDS
"""

import pyautogui
import time
import keyboard
import json

print("="*60)
print("Oracle界面坐标获取工具")
print("="*60)
print()
print("操作说明：")
print("  Ctrl+C  - 记录当前鼠标坐标")
print("  Ctrl+Q  - 退出并保存")
print()
print("请确保Oracle客户端已打开并显示在屏幕上")
print("等待5秒后开始...")
time.sleep(5)

coords = {}
labels = [
    ("import_button", "导入按钮"),
    ("file_input", "文件路径输入框"),
    ("confirm_button", "确认/上传按钮"),
    ("export_button", "导出销售订单按钮"),
    ("sales_order_field", "销售订单号显示区域"),
]

current_label_index = 0

def record_coord():
    global current_label_index
    if current_label_index < len(labels):
        key, desc = labels[current_label_index]
        x, y = pyautogui.position()
        coords[key] = (x, y)
        print(f"✅ 已记录 {desc}: ({x}, {y})")
        current_label_index += 1
        
        if current_label_index < len(labels):
            next_key, next_desc = labels[current_label_index]
            print(f"👉 下一步: 移动鼠标到【{next_desc}】，按Ctrl+C记录")
        else:
            print("✅ 所有坐标已记录完成！按Ctrl+Q退出并保存")
    else:
        x, y = pyautogui.position()
        coords[f"extra_{len(coords)}"] = (x, y)
        print(f"✅ 已记录额外坐标: ({x}, {y})")

def save_and_exit():
    print()
    print("="*60)
    print("记录的坐标：")
    print("="*60)
    
    # 输出到config格式的字符串
    config_str = "\nORACLE_COORDS = {\n"
    for key, value in coords.items():
        config_str += f"    \"{key.split('_extra')[0]}\": ({value[0]}, {value[1]}),\n"
    config_str += "}\n"
    
    print(config_str)
    
    # 保存到文件
    with open("oracle_coords_output.txt", "w", encoding="utf-8") as f:
        f.write("# Oracle界面坐标配置\n")
        f.write("# 请将以下内容复制到config.py的ORACLE_COORDS中\n\n")
        f.write(config_str)
    
    print("坐标已保存到 oracle_coords_output.txt")
    print("请将上面的坐标复制到 config.py 的 ORACLE_COORDS 配置中")
    exit(0)

# 监听按键
keyboard.add_hotkey('ctrl+c', record_coord)
keyboard.add_hotkey('ctrl+q', save_and_exit)

# 显示当前鼠标位置（实时更新）
print()
print(f"👉 请移动鼠标到【{labels[0][1]}】，按Ctrl+C记录")
print()
print("实时鼠标位置（用于参考）：")

try:
    while True:
        x, y = pyautogui.position()
        print(f"  当前位置: ({x}, {y})")
        time.sleep(1)
except KeyboardInterrupt:
    save_and_exit()
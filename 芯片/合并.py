import os
import pandas as pd

def merge_excel_files(folder_path, output_file="D://供应链//芯片//合并结果.xlsx"):
    """
    合并指定文件夹下的1-7.xlsx文件（表头相同）
    :param folder_path: 存放Excel文件的文件夹路径
    :param output_file: 合并后保存的文件名
    """
    # 存储所有文件的数据
    all_data = []
    
    # 遍历1-7号文件（根据实际文件名调整，例如“提取结果1.xlsx”到“提取结果7.xlsx”）
    for i in range(1, 12):  # 1到7
        # 构造文件名（根据实际文件名修改，例如"提取结果{}.xlsx"）
        file_name = f"D://供应链//芯片//提取结果{i}.xlsx"  # 若文件名为“1.xlsx”“2.xlsx”...用这个
        # 若文件名为“提取结果1.xlsx”，则改为：file_name = f"提取结果{i}.xlsx"
        
        # 拼接完整文件路径
        file_path = os.path.join(folder_path, file_name)
        print(file_path)
        # 检查文件是否存在
        if not os.path.exists(file_path):
            print(f"警告：文件不存在 - {file_path}，已跳过")
            continue
        
        try:
            # 读取Excel文件（表头已固定，无需额外处理）
            df = pd.read_excel(file_path, engine="openpyxl")
            # 检查表头是否符合预期（可选，用于验证文件正确性）
            expected_columns = ["第7列数据", "第8列数据", "第1行第11列数据"]
            if list(df.columns) != expected_columns:
                print(f"警告：文件{file_name}表头不匹配，已跳过")
                continue
            # 添加到总数据中
            all_data.append(df)
            print(f"已读取：{file_name}（{len(df)}行数据）")
        except Exception as e:
            print(f"读取文件{file_name}失败：{str(e)}，已跳过")
            continue
    
    # 合并所有数据
    if all_data:
        merged_df = pd.concat(all_data, ignore_index=True)  # ignore_index重置索引
        # 保存合并结果
        merged_df.to_excel(output_file, index=False, engine="openpyxl")
        print(f"\n合并完成！共{len(merged_df)}行数据，已保存到：{output_file}")
    else:
        print("\n未找到有效文件或所有文件读取失败，无法合并")

if __name__ == "__main__":
    # 1. 替换为你的Excel文件所在文件夹路径（例如："C:/data" 或 "./excel_files"）
    folder_path = "D://供应链//芯片//"  # 当前目录下的excel_files文件夹（可修改）
    
    # 2. 执行合并
    merge_excel_files(folder_path)
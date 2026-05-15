from lxml import etree
import pandas as pd
import os

def extract_table_data(html_content):
    """
    提取表格所有行的td[5]、td[6]、td[11]数据
    :param html_content: HTML内容字符串
    :return: 提取的数据列表（每行包含[td5, td6, td11]）
    """
    # 解析HTML
    parser = etree.HTMLParser()
    tree = etree.fromstring(html_content, parser)
    
    # 1. 获取表格总行数（tr的数量，索引从1开始）
    tr_xpath = '//*[@id="app"]/div/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div/div/div/div[2]/table/tbody/tr'
    all_tr = tree.xpath(tr_xpath)
    row_count = len(all_tr)
    if row_count == 0:
        print("未找到表格行数据，请检查XPath或HTML内容")
        return []
    
    # 2. 定义带行索引占位符的XPath（修正笔误，确保格式为//*开头）
    base_xpaths = {
        "td5": '//*[@id="app"]/div/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div/div/div/div[2]/table/tbody/tr[{i}]/td[5]/div/span',
        "td6": '//*[@id="app"]/div/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div/div/div/div[2]/table/tbody/tr[{i}]/td[6]/div/span',
        "td11": '//*[@id="app"]/div/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div/div/div/div[2]/table/tbody/tr[{i}]/td[11]/div/span'  # 修正XPath格式
    }
    
    # 3. 遍历所有行，提取数据
    all_data = []
    for i in range(1, row_count + 1):
        # 提取td5数据
        td5_path = base_xpaths["td5"].format(i=i)
        td5_elements = tree.xpath(td5_path)
        td5_text = td5_elements[0].text.strip() if (td5_elements and td5_elements[0].text) else ""
        
        # 提取td6数据
        td6_path = base_xpaths["td6"].format(i=i)
        td6_elements = tree.xpath(td6_path)
        td6_text = td6_elements[0].text.strip() if (td6_elements and td6_elements[0].text) else ""
        
        # 提取td11数据
        td11_path = base_xpaths["td11"].format(i=i)
        td11_elements = tree.xpath(td11_path)
        td11_text = td11_elements[0].text.strip() if (td11_elements and td11_elements[0].text) else ""
        
        # 添加当前行数据
        all_data.append([td5_text, td6_text, td11_text])
    
    return all_data

if __name__ == "__main__":
    # 1. 读取HTML内容（本地文件或网络请求）
    try:
        # 示例1：从本地HTML文件读取（需将文件命名为input.html放在同目录）
        with open("D://供应链//芯片//test.html", "r", encoding="utf-8") as f:
            html_content = f.read()
    except FileNotFoundError:
        print("未找到input.html，请将HTML文件放在代码同目录，或修改为网络请求")
        exit()
    
    # 示例2：从网络获取（需安装requests：pip install requests）
    # import requests
    # url = "目标网页URL"
    # response = requests.get(url)
    # html_content = response.text  # 部分网站可能需要指定编码，如response.content.decode('gbk')
    
    # 2. 提取数据
    table_data = extract_table_data(html_content)
    
    # 3. 保存为Excel
    if table_data:
        # 定义列名（可根据实际业务含义修改，如“名称”“数量”“金额”等）
        columns = ["第5列数据", "第6列数据", "第11列数据"]
        df = pd.DataFrame(table_data, columns=columns)
        
        # 保存路径（当前目录下的“提取结果.xlsx”）
        output_path = "D://供应链//芯片//提取结果.xlsx"
        df.to_excel(output_path, index=False, engine="openpyxl")
        print(f"成功提取{len(table_data)}行数据，已保存到：{os.path.abspath(output_path)}")
    else:
        print("未提取到有效数据，请检查XPath是否正确")
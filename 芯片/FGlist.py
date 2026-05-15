from lxml import etree
import pandas as pd

def extract_data(html_content):
    """
    提取所有行的td[7]、td[8]，以及第1行的td[11]数据
    :param html_content: HTML内容字符串
    :return: 整合后的数据列表（每行包含[td7, td8, td11_tr1]）
    """
    parser = etree.HTMLParser()
    tree = etree.fromstring(html_content, parser)
    
    # 1. 确定表格总行数（用于遍历tr[i]）
    tr_xpath = '/html/body/div[1]/div/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div/div/div/div[2]/table/tbody/tr'
    all_tr = tree.xpath(tr_xpath)
    row_count = len(all_tr)  # 总行数（索引从1开始）
    if row_count == 0:
        return []
    
    # 2. 提取固定行（tr[1]）的td[11]数据（仅提取一次）
    td11_tr1_xpath = '/html/body/div[1]/div/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div/div/div/div[2]/table/tbody/tr[1]/td[11]/div/span'
    td11_elements = tree.xpath(td11_tr1_xpath)
    td11_tr1_text = td11_elements[0].text.strip() if (td11_elements and td11_elements[0].text) else ""
    
    # 3. 遍历所有行，提取每行的td[7]和td[8]，并与td11_tr1组合
    all_data = []
    # 定义带占位符的XPath（{i}为行索引）
    base_xpath_td7 = '/html/body/div[1]/div/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div/div/div/div[2]/table/tbody/tr[{i}]/td[7]/div/span'
    base_xpath_td8 = '/html/body/div[1]/div/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div/div/div/div[2]/table/tbody/tr[{i}]/td[8]/div/span'
    
    for i in range(1, row_count + 1):
        # 提取当前行的td[7]
        xpath_td7 = base_xpath_td7.format(i=i)
        td7_elements = tree.xpath(xpath_td7)
        td7_text = td7_elements[0].text.strip() if (td7_elements and td7_elements[0].text) else ""
        
        # 提取当前行的td[8]
        xpath_td8 = base_xpath_td8.format(i=i)
        td8_elements = tree.xpath(xpath_td8)
        td8_text = td8_elements[0].text.strip() if (td8_elements and td8_elements[0].text) else ""
        
        # 每行数据包含：td7、td8、tr[1]的td11
        all_data.append([td7_text, td8_text, td11_tr1_text])
    
    return all_data

if __name__ == "__main__":
    # 1. 读取HTML内容（本地文件或网络请求）
    try:
        # 示例1：从本地HTML文件读取（需将文件命名为input.html放在同目录）
        with open("D://供应链//芯片//input.html", "r", encoding="utf-8") as f:
            html_content = f.read()
    except FileNotFoundError:
        print("未找到input.html，请将HTML文件放在代码同目录，或修改为网络请求")
        exit()
    
    # 示例2：从网络获取（需安装requests：pip install requests）
    # import requests
    # url = "目标网页URL"
    # response = requests.get(url)
    # html_content = response.text  # 注意：部分网站需处理编码（如response.content.decode('gbk')）
    
    # 2. 提取数据
    extracted_data = extract_data(html_content)
    
    # 3. 保存为Excel
    if extracted_data:
        # 定义列名（可根据实际业务含义修改，例如“数量”“金额”“备注”等）
        columns = ["第7列数据", "第8列数据", "第1行第11列数据"]
        df = pd.DataFrame(extracted_data, columns=columns)
        
        # 保存到Excel
        excel_path = "D://供应链//芯片//提取结果12.xlsx"
        df.to_excel(excel_path, index=False, engine="openpyxl")
        print(f"成功提取{len(extracted_data)}行数据，已保存到：{excel_path}")
    else:
        print("未提取到数据，请检查XPath是否正确或HTML内容是否包含目标表格")
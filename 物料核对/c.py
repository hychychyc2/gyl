from flask import Flask, request, render_template_string, send_file
import pandas as pd
import io
import datetime
import os
app = Flask(__name__)

# 网页模板
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html>
<head>
    <title>Excel表格处理工具</title>
    <style>
        body { font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; padding: 20px; }
        .container { text-align: center; }
        .upload-form { margin: 20px 0; padding: 20px; border: 1px solid #ccc; border-radius: 5px; }
        .file-input { margin: 10px 0; }
        input[type="submit"] { background-color: #4CAF50; color: white; padding: 10px 15px; border: none; border-radius: 4px; cursor: pointer; }
        input[type="submit"]:hover { background-color: #45a049; }
        .result { margin-top: 20px; padding: 10px; }
        .success { color: green; }
        .error { color: red; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Excel表格处理工具</h1>
        <div class="upload-form">
            <form method="POST" enctype="multipart/form-data">
                <h3>上传第一个Excel表格</h3>
                <input type="file" name="file1" accept=".xlsx, .xls" class="file-input" required><br>
                
                <h3>上传第二个Excel表格</h3>
                <input type="file" name="file2" accept=".xlsx, .xls" class="file-input" required><br><br>
                
                <input type="submit" value="处理表格">
            </form>
        </div>
        
        {% if message %}
        <div class="result {{ message_type }}">
            {{ message }}
        </div>
        {% endif %}
        
        {% if download_url %}
        <div class="download">
            <p>处理完成！点击下方链接下载结果：</p>
            <a href="{{ download_url }}" download>下载处理后的Excel文件</a>
        </div>
        {% endif %}
    </div>
</body>
</html>
'''



def read_html_file(file_path, skiprows=0):
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            html_content = f.read()
        dfs = pd.read_html(html_content, skiprows=skiprows)
        return dfs[0] if dfs else pd.DataFrame()
    except Exception as e:
        print(f"读取HTML文件错误: {str(e)}")
        raise
f={}
def find(tmp,x):
    f[tmp][x]=x if f[tmp][x]==x else find(tmp,f[tmp][x])
    return f[tmp][x]

bom_dict = {}
next_id={}
def remove_last_dot_and_after(s):
    # 找到最后一个点的位置
    last_dot_index = s.rfind('.')
    # 如果找到了点，则返回点之前的部分，否则返回原字符串
    if last_dot_index != -1:
        return s[:last_dot_index]
    return s

def process_first_excel(file):
    df = read_html_file(file)
    # import pdb;pdb.set_trace()
    # 初始化变量
    # last_id=0
    # for index, row in df.iterrows():
    #     if last_id!=row.iloc[0]:
    #         bom_dict[df[last_id].iloc[0]]=df[last_id:index]
    #     last_id=row.iloc[0]
    # bom_dict[df[last_id].iloc[0]]=df[last_id:index+1]
    for index, row in df.iterrows():
        if(row.iloc[0] not in next_id):
            next_id[row.iloc[0]]={}
            bom_dict[row.iloc[0]]={}
        is_alternative = str(row.iloc[4]).strip().endswith('替代')
        if(is_alternative):
            tmp_id=remove_last_dot_and_after(row.iloc[4])
        else:
            tmp_id=row.iloc[4]
        f_id=remove_last_dot_and_after(tmp_id)
        if(f_id not in next_id[row.iloc[0]]):
            next_id[row.iloc[0]][f_id]=[]
        if(tmp_id not in bom_dict[row.iloc[0]]):
            bom_dict[row.iloc[0]][tmp_id]=[]
        # if(row.iloc[0]=='C06010836'):
        #     print('/uu',row.iloc[4],f_id,tmp_id)
        if(tmp_id not in next_id[row.iloc[0]][f_id]):
            next_id[row.iloc[0]][f_id].append(tmp_id)
        # print("_______",row.iloc[8],type(row.iloc[8]))
        sunhao=0
        if(row.iloc[8]!="&nbsp"):
           sunhao=float(row.iloc[8])
        tmp_dict=(row.iloc[5],row.iloc[12],sunhao)
        bom_dict[row.iloc[0]][tmp_id].append(tmp_dict)
    return bom_dict

def process_first_excel_tmp(file):
    """处理第一个Excel表格，生成指定格式的字典"""
    # 读取Excel文件
    df = read_html_file(file)
    
    # 初始化变量
    result_dict = {}
    result_dict_pcb = {}

    main_ingredient = None  # 主料
    result_dict2 = {}
    result_dict2_pcb = {}
   
    sums={}
    sums_pcb={}
    tmp_pre=[]
    # 遍历表格行（忽略表头）
    for index, row in df.iterrows():
        # 获取当前行的第五列值（索引为4）
        is_alternative = str(row.iloc[4]).strip().endswith('替代')

        current_fifth_col = row.iloc[5] if len(row) > 5 else None
        
        # 更新主料
        if not is_alternative:
            main_ingredient = current_fifth_col
        
        # 获取所需列的值
        first_col = row.iloc[0] if len(row) > 0 else None
        sixth_col = row.iloc[5] if len(row) > 5 else None
        thirteenth_col = row.iloc[12] if len(row) > 12 else None
        if "PCB裸板" not in  row.iloc[6] and "控制板" in row.iloc[6]:
            tmp_pre.append(row.iloc[4])
            # print(tmp_pre)
        # 构建字典
        if first_col not in result_dict:
            result_dict[first_col] = {}
            result_dict_pcb[first_col] = {}
            
            f[first_col] = {}
        if sixth_col not in result_dict[first_col]:
            result_dict[first_col][sixth_col]={}
            result_dict_pcb[first_col][sixth_col]={}
        if main_ingredient in result_dict[first_col][sixth_col].keys():
            ff=0
            for i in tmp_pre:

                if(i not in row.iloc[4]):
                    ff=1
            if not ff:
                o2=result_dict[first_col][sixth_col][main_ingredient][0]
            else:
                o2=0
            o=result_dict[first_col][sixth_col][main_ingredient][0]
        else:
            o=0
            o2=0
        result_dict[first_col][sixth_col][main_ingredient] = (thirteenth_col+o, main_ingredient)
        result_dict_pcb[first_col][sixth_col][main_ingredient] = (thirteenth_col+o2, main_ingredient)

        if(main_ingredient not in f[first_col].keys()):
            f[first_col][main_ingredient]=main_ingredient
        f[first_col][sixth_col]=find(first_col,main_ingredient)
    print("+++++++++++")
    for i,j in result_dict.items():
        # print(i,j)
        if i not in sums:
            sums[i] = {}
        for k,p in j.items():
            main_ingredient=f[i][k]
            if(main_ingredient not in sums[i]):
                sums[i][main_ingredient]=0

            # print(p[k][0])

            if k in p and p[k][1]==k:
                # if(i=='C06010705' and main_ingredient=='Y02110411K01'):
                #     print(i,k,p[k])
                #     print(p[k][0])
                sums[i][main_ingredient] +=p[k][0]
    for i,j in result_dict_pcb.items():
        # print(i,j)
        if i not in sums_pcb:
            sums_pcb[i] = {}
        for k,p in j.items():
            main_ingredient=f[i][k]
            if(main_ingredient not in sums_pcb[i]):
                sums_pcb[i][main_ingredient]=0

            # print(p[k][0])

            if k in p and p[k][1]==k:
                # if(i=='C06010705' and main_ingredient=='Y02110411K01'):
                #     print(i,k,p[k])
                #     print(p[k][0])
                sums_pcb[i][main_ingredient] +=p[k][0]
    print("+++++++++++")

    for i,j in result_dict.items():
        if i not in result_dict2:
            result_dict2[i] = {}
        for k,p in j.items():
            main_ingredient=f[i][k]
            result_dict2[i][k]=(sums[i][main_ingredient],main_ingredient)
           
    for i,j in result_dict_pcb.items():
        if i not in result_dict2_pcb:
            result_dict2_pcb[i] = {}
        for k,p in j.items():
            main_ingredient=f[i][k]
            result_dict2_pcb[i][k]=(sums_pcb[i][main_ingredient],main_ingredient)
           
    # print(result_dict2)
    return (result_dict2,result_dict2_pcb)
tmp_mer={}
shaotou={}
def check(order_id,z_id,tot,id,fan):
    main_num=1
    tot_need=tot*main_num

    if(id in bom_dict[z_id]):
        main_num=bom_dict[z_id][id][0][1]
        # print("bom",bom_dict[z_id][id][0][1])
        
        tot_need=tot*main_num
        if(order_id=='PIEM20240201002-ZB' ):
            print(tot_need)
        if( ('（YZ）' in str(order_id) or'(YZ）' in str(order_id)or'（YZ)' in str(order_id) or'(YZ)' in str(order_id) )and tot <=100):
            # print((bom_dict[z_id][id][0][2]))
            tot_need=tot*main_num+float(bom_dict[z_id][id][0][2])
        # if(z_id=="C06010705" and id=='S.10.220'):
        #     print("HH",bom_dict[z_id][id])
        for i in bom_dict[z_id][id]:
            # print(i,tmp_mer[order_id])
            if(order_id=='PIEM20240201002-ZB'):
                print(id,i,tot_need)
            if(i[0] in tmp_mer[order_id]):
                if(order_id=='PIEM20240201002-ZB' and (i[0]=='Y02110511K01'or 0)):
                    print(order_id,z_id,tot,id)

                    print(z_id,id)
                    print(order_id)
                    print("bom",bom_dict[z_id][id][0][1])
                    print("bom2",bom_dict[z_id][id])

                    print(i,tmp_mer[order_id])
                    print("oder",tmp_mer[order_id][i[0]],i[1])
                    print(tot_need)
                # print(i[1],bom_dict[z_id][id][0][1])
                if(i[1] and tmp_mer[order_id][i[0]]/abs(i[1])*bom_dict[z_id][id][0][1]>=tot_need):
                    tmp_mer[order_id][i[0]]-=tot_need/bom_dict[z_id][id][0][1]*i[1]
                    tot_need=0
                    if(order_id=='PIEM20240201002-ZB' and (i[0]=='Y02110511K01'or 0)):
                        print(tmp_mer[order_id][i[0]])
                    return 1
                elif i[1] and pd.notnull(tmp_mer[order_id][i[0]]):
                    
                    tot_need-=tmp_mer[order_id][i[0]]/abs(i[1])*bom_dict[z_id][id][0][1]
                    tmp_mer[order_id][i[0]]=0

                    if(order_id=='PIEM20240201002-ZB' and (i[0]=='Y02110511K01'or 0)):
                        print(tmp_mer[order_id][i[0]],tot_need)
                
    ok=tot_need==0
    # print(next_id[z_id])
    if(id in next_id[z_id]and not ok):
        ok=1 
        # print(z_id)
        # if(z_id=="C06010836"):
        #     print("uuuu",id)
        #     print(next_id[z_id][id])
        for i in next_id[z_id][id]:
            # if(order_id=='YNAH20250901009-K1' and i=='S.10'):
            #     print("uuuu",id)
            #     print(next_id[z_id][id])
            #     print(next_id[z_id])
            #     if(i in next_id[z_id]):
            #         print(next_id[z_id][i])
            #     print("jj",i)
            #     print(order_id,z_id,tot_need/main_num,i)
            ok&=check(order_id,z_id,tot_need/main_num,i,fan)
    if not fan and not ok:
            if  (id in bom_dict[z_id]):
                if bom_dict[z_id][id][0][0] not in shaotou[order_id]:
                    shaotou[order_id][bom_dict[z_id][id][0][0]]=tot_need
                else:   
                    shaotou[order_id][bom_dict[z_id][id][0][0]]+=tot_need
            elif id=="S":
                shaotou[order_id][z_id]=tot_need

            # print(z_id,tot_need)
            # shaotou[order_id].append({bom_dict[z_id][id][0][0]:tot_need})
    return ok
def process_second_excel(file, first_dict):
    """处理第二个Excel表格，这里实现了基本框架，您可以根据需要补充具体处理逻辑"""
    global shaotou
    # 读取Excel文件
    df = pd.read_excel(file)
    
    # 这里可以根据第一个表格生成的字典对第二个表格进行处理
    # 示例：遍历第二张表的每一行，可以使用first_dict中的数据进行操作
    # 注意：请根据实际需求修改以下处理逻辑
    tmp={}
    tot={}
    shaotou={}
    # import pdb;pdb.set_trace()
    is_processd={}
    for index, row in df.iterrows():
        if row.iloc[0] not in tmp_mer:

            tmp_mer[row.iloc[0]]={}
            tot[row.iloc[0]]=row.iloc[3]
        if row.iloc[12] not in tmp_mer[row.iloc[0]]:
            tmp_mer[row.iloc[0]][row.iloc[12]]=0
        if pd.notnull(row.iloc[17]):
            tmp_mer[row.iloc[0]][row.iloc[12]]+=row.iloc[17]
    print("---------")
    for index, row in df.iterrows():
        if row.iloc[1] not in bom_dict:continue

        if(row.iloc[0][-2:]=='-B' or row.iloc[0][-6:]=='-B(YZ)'  or row.iloc[0][-6:]=='-B(YZ）' or row.iloc[0][-7:]=='-B1(YZ)' or row.iloc[0][-4:]=='(CG)' or row.iloc[0][-4:]=='(CG）'or row.iloc[0][-4:]=='（CG）'or row.iloc[0][-4:]=='（CG)' or row.iloc[0][-6:]=='(YZCG)'):
            if(row.iloc[1] == row.iloc[12]):
                tmp_mer[row.iloc[0]][row.iloc[12]]=0
            if tmp_mer[row.iloc[0]][row.iloc[12]]>0 and row.iloc[0] not in is_processd:
                is_processd[row.iloc[0]]=1
                check(row.iloc[0],row.iloc[1],row.iloc[3],'S',1)
            continue
        if row.iloc[0] in is_processd:continue
        if row.iloc[1] not in bom_dict:continue
        # print(row.iloc[13])
       
        # print(row.iloc[0],row.iloc[1],row.iloc[3])
        # print(bom_dict[row.iloc[1]])
        is_processd[row.iloc[0]]=1
        shaotou[row.iloc[0]]={}
        check(row.iloc[0],row.iloc[1],row.iloc[3],'S',0)
        if(row.iloc[0]=='PIEM20250601011-ZB'):
            print(shaotou[row.iloc[0]])

    print("---------2")
    # print(tmp_mer['MOLT20240801001-ZB']['YBL0B0202064'])
    

    for index, row in df.iterrows():
        if(row.iloc[0]=='NGSB20250501002-ZB' and (row.iloc[12]=='Y01101041F00'or 0)):

            print(tmp_mer[row.iloc[0]][row.iloc[12]])
            print(df.iloc[index,23])
        if pd.notnull(df.iloc[index,23]):
            df.iloc[index,23]=min(max(df.iloc[index,23],0),tmp_mer[row.iloc[0]][row.iloc[12]])
        else:
            df.iloc[index,23]=tmp_mer[row.iloc[0]][row.iloc[12]]
    data=[]
    print("---------3")

    for i,j in shaotou.items():
        for k,p in j.items():
            data.append({' 工单编号':i,'组件编码':k,'需求量差':-p,'已投料差':0})
    new_rows = pd.DataFrame(
        data=data,
        columns=df.columns  # 强制使用原DataFrame的列结构（关键！）
    )
    df_with_new = pd.concat([df, new_rows], ignore_index=True)
    print("---------4")

    # 2. 按第一列排序（自动处理空值，默认NaN放最后）
    # 第一列列名通过 df.columns[0] 获取，无需硬编码（通用关键！）
    df = df_with_new.sort_values(
        by=df.columns[0],
        na_position='last'  # 可选：'last'（默认）或'first'（空值放前面）
    )
    print(df.columns)
    df.insert(25, '总差异', df['需求量差'] + df['已投料差'])
    df.insert(26, '比特物控回复', "OK,可关闭")  
    import numpy as np
    df['比特物控回复'] = np.where(
    df['已投料差'] != 0,  # 第一个条件：Y不等于0
    "投料有差异，请完成投料",
  
      
        np.where(
        # 确保所有条件都是布尔数组，用括号明确分组
        (df['需求量差'] != -9999999) &  # 布尔数组
        (
            (abs(df['需求量差']) < 0.99999)  # 布尔数组
            | 
            (
                (df['需求量差'] <= 20)  # 布尔数组
                & 
                (
                    df['组件描述'].str.contains("电阻", na=False)  # 布尔数组
                    | df['组件描述'].str.contains("电容", na=False)
                    | df['组件描述'].str.contains("电感", na=False)
                    | df['组件描述'].str.lower().str.contains("res", na=False)
                    | df['组件描述'].str.lower().str.contains("smd", na=False)
                )
            )
        ),
        "OK,可关闭",
        np.where(
            df['组件描述'].str.contains("加工费", na=False),  # 布尔数组
            "加工费不核对",
            "需求差异"
        )
    )
    
)

    print("+++++++++++")

    # for index, row in df.iterrows():
    #     if(row.iloc[0] not in tmp.keys()):
    #         df.iloc[index, 23] = 0           

    #         continue
    #     if(row.iloc[12] in first_dict[is_pcba[row.iloc[0]]][row.iloc[1]].keys() and first_dict[is_pcba[row.iloc[0]]][row.iloc[1]][row.iloc[12]][1] in tmp[row.iloc[0]].keys()):
    #         df.iloc[index, 23]=tmp[row.iloc[0]][first_dict[is_pcba[row.iloc[0]]][row.iloc[1]][row.iloc[12]][1]]
    #         if(df.iloc[index, 23]==0 and(row.iloc[0][-2:]=='-B' or row.iloc[0][-6:]=='-B(YZ)'  or row.iloc[0][-6:]=='-B(YZ）' or row.iloc[0][-7:]=='-B1(YZ)' or row.iloc[0][-4:]=='(CG)' or row.iloc[0][-6:]=='(YZCG)')):
    #             df.iloc[index, 20]=0
    #         else:
    #             main_metra=row.iloc[12]
    #             while(first_dict[is_pcba[row.iloc[0]]][row.iloc[1]][main_metra][1]!=main_metra and main_metra):
    #                 main_metra=first_dict[is_pcba[row.iloc[0]]][row.iloc[1]][main_metra][1]
    #             df.iloc[index, 20]=first_dict[is_pcba[row.iloc[0]]][row.iloc[1]][main_metra][0]*row.iloc[3]
    #         df.iloc[index, 23]=df.iloc[index, 23]-df.iloc[index, 20]
    #         df.iloc[index, 11]=first_dict[is_pcba[row.iloc[0]]][row.iloc[1]][row.iloc[12]][1]    

    #     else:
    #         if(row.iloc[1]==row.iloc[12]):
    #             df.iloc[index, 23]=df.iloc[index, 17]-df.iloc[index, 3]
    #         else:
    #             df.iloc[index, 23]=-9999999
        # 示例操作：可以根据实际需求修改
    # 删除第12列（索引为17）值为零的行
    # for index, row in df.iterrows():
    #     if(row.iloc[0] not in tmp.keys()):
    #         print(row.iloc[23])
    # if len(df.columns) > 23:
    #     df = df[df.iloc[:, 23] != 0]
    
    return df
UPLOAD_FOLDER = 'uploads'

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # 检查文件是否上传
        if 'file1' not in request.files or 'file2' not in request.files:
            return render_template_string(HTML_TEMPLATE, 
                                        message="请上传两个Excel文件", 
                                        message_type="error")
        
        file1 = request.files['file1']
        file2 = request.files['file2']
        
        # 检查文件是否为空
        if file1.filename == '' or file2.filename == '':
            return render_template_string(HTML_TEMPLATE, 
                                        message="请选择有效的文件", 
                                        message_type="error")
        
        try:
            # 处理第一个Excel表格
            xlsx_path = os.path.join(UPLOAD_FOLDER, f"_1.{file1.filename.split('.')[-1]}")

            file1.save(xlsx_path)

            print(file1)
            first_dict= process_first_excel(xlsx_path)
            
            # 重置文件指针，因为已经读取过一次
            file2.seek(0)
            
            # 处理第二个Excel表格
            processed_df = process_second_excel(file2, first_dict)
            
            # 将处理后的DataFrame保存到内存中
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                processed_df.to_excel(writer, index=False)
            
            # 重置输出流指针
            output.seek(0)
            
            # 生成唯一的文件名
            timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
            filename = f"processed_file_{timestamp}.xlsx"
            
            # 返回处理后的文件
            return send_file(output, 
                            download_name=filename,
                            as_attachment=True,
                            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
        except Exception as e:
            return render_template_string(HTML_TEMPLATE, 
                                        message=f"处理文件时出错: {str(e)}", 
                                        message_type="error")
    
    # GET请求，显示上传表单
    return render_template_string(HTML_TEMPLATE)

if __name__ == '__main__':
    app.run(debug=True)
# # import os
# # from datetime import datetime
# # from flask import Flask, render_template, request, send_file
# # from openpyxl import load_workbook
# # import pandas as pd
# # # import xlwings as xw
# # app = Flask(__name__)
# # UPLOAD_FOLDER = 'uploads'
# # OUTPUT_FOLDER = 'outputs'
# # ALLOWED_EXTENSIONS = {'xlsx', 'xlsm'}

# # os.makedirs(UPLOAD_FOLDER, exist_ok=True)
# # os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# # def allowed_file(filename):
# #     return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS
# # import mimetypes as py_mimetypes
# # def process_excel(xlsx_path, xlsm_path):

# #     py_mimetypes.types_map['.JPG'] = 'image/jpeg'
# #     current_month = datetime.now().strftime('%Y%m')
# #     current_day = datetime.now().strftime('%Y%m%d')

# #     subject = "东莞群光"  # 这里需要根据实际主体信息修改
    
# #     # 读取xlsx数据
# #     xlsx_df = pd.read_excel(xlsx_path)
# #     xlsx_data = xlsx_df.to_dict('records')
    
# #     # 加载xlsm文件，保留宏
# #     wb = load_workbook(xlsm_path, data_only=False, keep_vba=True)
# #     sheet = wb.active
    
# #     # 假设xlsm表头在第3行（前两行空，第一列空），实际需根据真实结构调整
# #     header_row = 3
# #     data_start_row = header_row + 1
    
# #     # 匹配列索引（需根据实际表头调整）
# #     col_mapping = {
# #         '订单编号': None,
# #         '订单日期': None,
# #         '发运说明': None,
# #         '客户PO': None,
# #         '物料编码': None,
# #         '数量': None
# #     }
# #     for col_idx, cells in enumerate(sheet[header_row], 1):
# #         if cells.value in col_mapping:
# #             col_mapping[cells.value] = col_idx - 1  # 转换为0-based索引
# #     # 清理数据行（假设从data_start_row开始，删除多余行）
# #     # 这里需要根据实际数据行数处理，示例中假设xlsx和xlsm数据行数一致
# #     row_idx2= data_start_row+1
# #     max_row = sheet.max_row
# #     for row_idx, data in enumerate(xlsx_data, data_start_row):
# #         # 订单编号替换为销售订单
# #         if(data['外协']!=subject):continue

# #         sheet.cells(row=row_idx2, column=col_mapping['订单编号']+1).value = str(data['销售订单'])
# #         # print(data['销售订单'])
# #         # print(sheet.cells(row=row_idx2, column=col_mapping['订单编号']+1).value)

# #         # 订单日期改成本月（假设格式为yyyymmdd）
# #         sheet.cells(row=row_idx2, column=col_mapping['订单日期']+1).value = current_day
# #         # 发运说明和客户PO添加主体和本月超领
# #         sheet.cells(row=row_idx2, column=col_mapping['发运说明']+1).value = f"{subject}{current_month}超领"
# #         sheet.cells(row=row_idx2, column=col_mapping['客户PO']+1).value = f"{subject}{current_month}超领"
# #         # 物料编码替换为物料号
# #         sheet.cells(row=row_idx2, column=col_mapping['物料编码']+1).value = data['物料号']
# #         # 数量替换为超领申请数量
# #         sheet.cells(row=row_idx2, column=col_mapping['数量']+1).value = data['超领申请数量']
# #         row_idx2+=1
# #         line_num=max_row-row_idx2
# #     sheet.delete_rows(row_idx2,line_num)
# #     # 删除多余行（示例中假设只保留数据行，需根据实际逻辑调整）
# #     # 这里暂不处理复杂删除逻辑，实际需根据业务需求完善
# #     # print(wb)
# #     # print(sheet)
# #     output_path = os.path.join(OUTPUT_FOLDER, 'processed_file.xlsm')
# #     for img in sheet._images:
# #         img_path = img._path  # 获取图片在 Excel 中的路径（如 'xl/media/image1.JPG'）
# #         print(img_path)
# #         ext = os.path.splitext(img_path)[1]  # 获取扩展名（'.JPG'）
# #         if ext.upper() == '.JPG':
# #             # 将扩展名转为小写 .jpg
# #             new_img_path = img_path[:-4] + '.jpg'
# #             img._path = new_img_path  # 更新图片路径
# #     wb.save(output_path)
# #     return output_path

# # @app.route('/', methods=['GET', 'POST'])
# # def upload_file():
# #     if request.method == 'POST':
# #         xlsx_file = request.files.get('xlsx_file')
# #         xlsm_file = request.files.get('xlsm_file')
        
# #         if not xlsx_file or not xlsm_file:
# #             return "请上传两个文件", 400
# #         if not (allowed_file(xlsx_file.filename) and allowed_file(xlsm_file.filename)):
# #             return "文件类型不允许", 400
        
# #         xlsx_path = os.path.join(UPLOAD_FOLDER, xlsx_file.filename)
# #         xlsm_path = os.path.join(UPLOAD_FOLDER, xlsm_file.filename)
# #         xlsx_file.save(xlsx_path)
# #         xlsm_file.save(xlsm_path)
        
# #         #try:
# #         processed_path = process_excel(xlsx_path, xlsm_path)
# #             # return send_file(processed_path, as_attachment=True)
# #         # except Exception as e:
# #         #     return f"处理失败: {str(e)}", 500
# #         # finally:
# #         #     os.remove(xlsx_path)
# #         #     os.remove(xlsm_path)
    
# #     return render_template('index.html')

# # if __name__ == '__main__':
# #     app.run(debug=True)



# # import os
# # from datetime import datetime
# # from flask import Flask, render_template, request, send_file
# # import pandas as pd
# # import xlwings as xw  # 使用xlwings处理含宏的xlsm文件

# # app = Flask(__name__)
# # UPLOAD_FOLDER = 'uploads'
# # OUTPUT_FOLDER = 'outputs'
# # ALLOWED_EXTENSIONS = {'xlsx', 'xlsm'}

# # os.makedirs(UPLOAD_FOLDER, exist_ok=True)
# # os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# # def allowed_file(filename):
# #     return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# # def process_excel(xlsx_path, xlsm_path):
# #     current_month = datetime.now().strftime('%Y%m')
# #     current_day = datetime.now().strftime('%Y%m%d')
# #     subject = "东莞群光"  # 这里需要根据实际主体信息修改

# #     # 读取xlsx数据
# #     xlsx_df = pd.read_excel(xlsx_path)
# #     xlsx_data = xlsx_df.to_dict('records')

# #     # 使用xlwings打开xlsm文件（保留宏）
# #     wb = xw.Book(xlsm_path)
# #     sheet = wb.sheets[0]  # 假设操作第一个工作表

# #     # 定义列映射（根据实际表头在第3行，第一列从A开始）
# #     header_row = 3  # xlwings行号从1开始
# #     col_mapping = {
# #         '订单编号': None,
# #         '订单日期': None,
# #         '发运说明': None,
# #         '客户PO': None,
# #         '物料编码': None,
# #         '数量': None
# #     }

# #     # 查找列索引
# #     for col_idx, cells_value in enumerate(sheet.range((header_row, 1), (header_row, sheet.api.UsedRange.Columns.Count)).value, 1):
# #         if cells_value in col_mapping:
# #             col_mapping[cells_value] = col_idx  # xlwings列号从1开始

# #     # 数据起始行（假设从第4行开始）
# #     data_start_row = header_row + 1
# #     current_row = data_start_row

# #     # 清理原有数据行（从数据起始行开始删除所有行）
# #     # if sheet.api.UsedRange.Rows.Count > data_start_row:
# #     #     sheet.api.Rows(f"{data_start_row}:{sheet.api.UsedRange.Rows.Count}").Delete()

# #     for data in xlsx_data:
# #         if data['外协'] != subject:
# #             continue

# #         # 订单编号
# #         sheet.range((current_row, col_mapping['订单编号'])).value = str(data['销售订单'])
# #         # 订单日期
# #         sheet.range((current_row, col_mapping['订单日期'])).value = current_day
# #         # 发运说明
# #         sheet.range((current_row, col_mapping['发运说明'])).value = f"{subject}{current_month}超领"
# #         # 客户PO
# #         sheet.range((current_row, col_mapping['客户PO'])).value = f"{subject}{current_month}超领"
# #         # 物料编码
# #         sheet.range((current_row, col_mapping['物料编码'])).value = data['物料号']
# #         # 数量
# #         sheet.range((current_row, col_mapping['数量'])).value = data['超领申请数量']

# #         current_row += 1

# #     # 保存文件（xlwings会正确保留VBA宏）
# #     output_path = os.path.join(OUTPUT_FOLDER, 'processed_file.xlsm')
# #     wb.save(output_path)
# #     wb.close()
# #     return output_path

# # @app.route('/', methods=['GET', 'POST'])
# # def upload_file():
# #     if request.method == 'POST':
# #         xlsx_file = request.files.get('xlsx_file')
# #         xlsm_file = request.files.get('xlsm_file')
        
# #         if not xlsx_file or not xlsm_file:
# #             return "请上传两个文件", 400
# #         if not (allowed_file(xlsx_file.filename) and allowed_file(xlsm_file.filename)):
# #             return "文件类型不允许", 400
        
# #         xlsx_path = os.path.join(UPLOAD_FOLDER, xlsx_file.filename)
# #         xlsm_path = os.path.join(UPLOAD_FOLDER, xlsm_file.filename)
# #         xlsx_file.save(xlsx_path)
# #         xlsm_file.save(xlsm_path)
        
# #         #try:
# #         processed_path = process_excel(xlsx_path, xlsm_path)
# #             # return send_file(processed_path, as_attachment=True)
# #         # except Exception as e:
# #         #     return f"处理失败: {str(e)}", 500
# #         # finally:
# #         #     os.remove(xlsx_path)
# #         #     os.remove(xlsm_path)
    
# #     return render_template('index.html')

# # if __name__ == '__main__':
# #     app.run(debug=True)
    


# import os
# from datetime import datetime
# from flask import Flask, render_template, request, send_file
# import pandas as pd
# import xlwings as xw

# app = Flask(__name__)
# UPLOAD_FOLDER = 'uploads'
# OUTPUT_FOLDER = 'outputs'
# ALLOWED_EXTENSIONS = {'xlsx', 'xlsm'}

# os.makedirs(UPLOAD_FOLDER, exist_ok=True)
# os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# def allowed_file(filename):
#     return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# def process_excel(xlsx_path, xlsm_path):
#     current_month = datetime.now().strftime('%Y%m')
#     current_day = datetime.now().strftime('%Y%m%d')
#     subject =  xlsm_path.rsplit('\\', 1)[-1].rsplit('.', 1)[0]
#     print(subject)
#     # 读取xlsx数据
#     xlsx_df = pd.read_excel(xlsx_path, dtype={'销售订单': str})
#     xlsx_data = xlsx_df.to_dict('records')
#     # 使用xlwings打开xlsm文件，保留宏
#     wb = xw.Book(xlsm_path)
#     sheet = wb.sheets['WebADI']  # 假设使用第一个工作表

#     # 定义表头行和数据起始行（根据实际情况调整）
#     header_row = 3
#     data_start_row = header_row + 2

#     # 构建列索引映射（根据实际表头获取列号）
#     col_mapping = {
#         '订单编号': None,
#         '订单日期': None,
#         '发运说明': None,
#         '客户PO': None,
#         '物料编码': None,
#         '数量': None
#     }
#     for col_idx in range(1,sheet.api.Columns.Count):
#         header_value = sheet.cells(header_row, col_idx + 1).value
#         # print(header_value)
#         if(header_value==None):break
#         if header_value in col_mapping:
#             col_mapping[header_value] = col_idx + 1  # xlwings列号从1开始
#     # 准备数据行操作
#     row_idx2 = data_start_row
#     max_row = sheet.api.UsedRange.Rows.Count  # 获取实际数据最大行


#     for data in xlsx_data:
#         if data['外协'] != subject:
#             if data['外协'] == None:
#                 break
#         # 填充订单编号
#             continue
#         if(row_idx2>data_start_row):
#             for j in range(2,23):
#                 if(sheet.cells(row_idx2, j).value == None):
#                     sheet.cells(row_idx2, j).value =  sheet.cells(row_idx2-1, j).value 
#         sheet.cells(row_idx2, col_mapping['订单编号']).value = f'="{data['销售订单']}"'

#         # 填充订单日期
#         sheet.cells(row_idx2, col_mapping['订单日期']).value = current_day
#         # 填充发运说明和客户PO
#         sheet.cells(row_idx2, col_mapping['发运说明']).value = f"{subject}{current_month}超领"
#         sheet.cells(row_idx2, col_mapping['客户PO']).value = f"{subject}{current_month}超领"
#         # 填充物料编码和数量
#         sheet.cells(row_idx2, col_mapping['物料编码']).value = data['物料号']
#         sheet.cells(row_idx2, col_mapping['数量']).value = data['超领申请数量']

       

#         row_idx2 += 1

#     # 删除多余行
#     if row_idx2 <= max_row:
#         sheet.api.Rows(f"{row_idx2}:{max_row}").Delete()

#     # 保存文件（xlwings会正确保留VBA宏）
#     output_path = os.path.join(OUTPUT_FOLDER, subject+'.xlsm')
#     wb.save(output_path)
#     wb.close()
#     return output_path

# @app.route('/', methods=['GET', 'POST'])
# def upload_file():
#     if request.method == 'POST':
#         xlsx_file = request.files.get('xlsx_file')
#         xlsm_file = request.files.get('xlsm_file')

#         if not xlsx_file or not xlsm_file:
#             return "请上传两个文件", 400
#         if not (allowed_file(xlsx_file.filename) and allowed_file(xlsm_file.filename)):
#             return "文件类型不允许", 400

#         xlsx_path = os.path.join(UPLOAD_FOLDER, xlsx_file.filename)
#         xlsm_path = os.path.join(UPLOAD_FOLDER, xlsm_file.filename)
#         xlsx_file.save(xlsx_path)
#         xlsm_file.save(xlsm_path)

#         #try:
#         processed_path = process_excel(xlsx_path, xlsm_path)
#         #    return send_file(processed_path, as_attachment=True)
#         # except Exception as e:
#         #     return f"处理失败: {str(e)}", 500
#         # finally:
#         #     os.remove(xlsx_path)
#         #     os.remove(xlsm_path)

#     return render_template('index.html')

# if __name__ == '__main__':
#     app.run(debug=True)



import os
from datetime import datetime
from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
import xlwings as xw
import threading
import uuid
from queue import Queue

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
ALLOWED_EXTENSIONS = {'xlsx', 'xlsm'}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# 用于存储任务状态（线程安全队列）
task_queue = Queue()
task_status = {}  # {task_id: {'status': 'pending', 'progress': 0, 'error': None, 'result': None}}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def process_excel_worker(task_id, xlsx_path, xlsm_path):
    try:
        task_status[task_id]['status'] = 'processing'
        current_month = datetime.now().strftime('%Y%m')
        current_day = datetime.now().strftime('%Y%m%d')
        subject = xlsm_path.rsplit('\\', 1)[-1].rsplit('.', 1)[0]
        xlsx_df = pd.read_excel(xlsx_path, dtype={'销售订单': str})
        xlsx_data = xlsx_df#.to_dict('records')
        xlsx_data = xlsx_data[xlsx_data['外协'] == subject]
        print(xlsx_data['外协'] == subject)
        total_rows = len(xlsx_data[xlsx_data['外协'] == subject])
        print(total_rows)
        xlsx_data = xlsx_data.to_dict('records')
        print(xlsx_data)
        wb = xw.Book(xlsm_path)
        sheet = wb.sheets['WebADI']
        header_row = 3
        data_start_row = header_row + 2
        
        col_mapping = {}
        for col_idx in range(1,sheet.api.Columns.Count):
            header_value = sheet.cells(header_row, col_idx + 1).value
            if(header_value==None):break
            if header_value in ['订单编号', '订单日期', '发运说明', '客户PO', '物料编码', '数量']:
                col_mapping[header_value] = col_idx + 1
        
        row_idx2 = data_start_row
        for i, data in enumerate(xlsx_data, 1):
            # if data['外协'] != subject:
            #     if data['外协'] == None:
            #         break
            # # 填充订单编号
            #     continue
            print(data['外协'])
            # 更新进度（每处理10%触发一次更新）
            progress = int((i / total_rows) * 100)
            if progress > task_status[task_id]['progress']:
                task_status[task_id]['progress'] = progress
            
            # 业务逻辑...（保持原有处理逻辑）
            if row_idx2 > data_start_row:
                for j in range(2, 23):
                    if sheet.cells(row_idx2, j).value is None:
                        sheet.cells(row_idx2, j).value = sheet.cells(row_idx2-1, j).value 
            sheet.cells(row_idx2, col_mapping['订单编号']).value = f'="{data["销售订单"]}"'
            sheet.cells(row_idx2, col_mapping['订单日期']).value = current_day
            sheet.cells(row_idx2, col_mapping['发运说明']).value = f"{subject}{current_month}超领"
            sheet.cells(row_idx2, col_mapping['客户PO']).value = f"{subject}{current_month}超领"
            sheet.cells(row_idx2, col_mapping['物料编码']).value = data['物料号']
            sheet.cells(row_idx2, col_mapping['数量']).value = data['超领申请数量']
            row_idx2 += 1
        
        # 删除多余行
        if row_idx2 <= sheet.api.UsedRange.Rows.Count:
            sheet.api.Rows(f"{row_idx2}:{sheet.api.UsedRange.Rows.Count}").Delete()
        
        output_path = os.path.join(OUTPUT_FOLDER, subject + '.xlsm')
        print(output_path)
        wb.save(output_path)
        wb.close()
        
        task_status[task_id]['status'] = 'complete'
        task_status[task_id]['result'] = output_path
    except Exception as e:
        task_status[task_id]['status'] = 'error'
        task_status[task_id]['error'] = str(e)
    # finally:
    #     # 清理临时文件
    #     os.remove(xlsx_path)
    #     os.remove(xlsm_path)

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        xlsx_file = request.files.get('xlsx_file')
        xlsm_file = request.files.get('xlsm_file')
        
        if not xlsx_file or not xlsm_file:
            return "请上传两个文件", 400
        if not (allowed_file(xlsx_file.filename) and allowed_file(xlsm_file.filename)):
            return "文件类型不允许", 400
        
        # 生成任务ID
        task_id = str(uuid.uuid4())
        xlsx_path = os.path.join(UPLOAD_FOLDER, xlsx_file.filename)
        xlsm_path = os.path.join(UPLOAD_FOLDER, xlsm_file.filename)
        xlsx_file.save(xlsx_path)
        xlsm_file.save(xlsm_path)
        
        # 初始化任务状态
        task_status[task_id] = {
            'status': 'pending',
            'progress': 0,
            'error': None,
            'result': None
        }
        
        # 启动后台线程处理任务
        threading.Thread(
            target=process_excel_worker,
            args=(task_id, xlsx_path, xlsm_path),
            daemon=True
        ).start()
        
        return jsonify({'task_id': task_id}), 202  # 返回任务ID
    
    return render_template('index.html')

@app.route('/get_progress/<task_id>')
def get_progress(task_id):
    status = task_status.get(task_id, {'status': 'not_found'})
    return jsonify(status)

@app.route('/download/<task_id>')
def download_file(task_id):
    status = task_status.get(task_id)
    if status and status['status'] == 'complete':
        output_path = status['result']
        return send_file(output_path, as_attachment=True)
    return "文件处理未完成", 404


if __name__ == '__main__':
     app.run(debug=True)
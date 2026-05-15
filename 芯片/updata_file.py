# import os
# import json
# import base64
# import requests
# import mimetypes  # 引入Python标准库的MIME类型识别模块

# def load_confluence_config(config_file_path):
#     """
#     读取JSON配置文件并校验关键配置项
#     :param config_file_path: JSON配置文件路径
#     :return: 解析后的配置字典
#     """
#     # 检查配置文件是否存在
#     if not os.path.exists(config_file_path):
#         raise FileNotFoundError(f"配置文件不存在：{config_file_path}")
    
#     # 读取并解析JSON
#     with open(config_file_path, 'r', encoding='utf-8') as f:
#         config = json.load(f)
    
#     # 校验核心配置项
#     required_keys = ['confluence_url', 'page_id', 'upload_config']
#     for key in required_keys:
#         if key not in config:
#             raise ValueError(f"配置文件缺少关键项：{key}")
    
#     # 校验认证信息（Cloud/Server二选一）
#     is_cloud = "atlassian.net" in config['confluence_url']
#     if is_cloud:
#         if not config['cloud_auth']['email'] or not config['cloud_auth']['api_token']:
#             raise ValueError("Cloud版配置缺少email或api_token")
#     else:
#         if not config['server_auth']['username'] or not config['server_auth']['password']:
#             raise ValueError("Server版配置缺少username或password")
    
#     # 校验文件路径
#     if not config['upload_config']['file_paths']:
#         raise ValueError("配置文件中未指定要上传的文件路径")
#     for file_path in config['upload_config']['file_paths']:
#         if not os.path.exists(file_path):
#             raise FileNotFoundError(f"指定的文件不存在：{file_path}")
    
#     return config

# def confluence_upload_from_config(config):
#     """
#     根据JSON配置上传文件到Confluence
#     :param config: 解析后的配置字典
#     :return: 上传结果（JSON）
#     """
#     # 提取基础配置
#     confluence_url = config['confluence_url']
#     page_id = config['page_id']
#     file_paths = config['upload_config']['file_paths']
#     comment = config['upload_config']['comment']
#     minor_edit = config['upload_config']['minor_edit']
    
#     # 构建认证头部
#     is_cloud = "atlassian.net" in confluence_url
#     if is_cloud:
#         auth_str = f"{config['cloud_auth']['email']}:{config['cloud_auth']['api_token']}"
#     else:
#         auth_str = f"{config['server_auth']['username']}:{config['server_auth']['password']}"
    
#     # 编码认证信息
#     auth_header = f"Basic {base64.b64encode(auth_str.encode()).decode()}"
#     headers = {
#         "Authorization": auth_header,
#         "X-Atlassian-Token": "no-check"  # 必须的CSRF防护头部
#     }

#     # 准备上传的文件列表
#     files = []
#     for file_path in file_paths:
#         file_name = os.path.basename(file_path)
#         # 修复：改用Python标准库mimetypes识别MIME类型（兼容所有版本）
#         mime_type, _ = mimetypes.guess_type(file_path)
#         if mime_type is None:
#             mime_type = "application/octet-stream"  # 通用二进制类型兜底
#         files.append(("file", (file_name, open(file_path, "rb"), mime_type)))
    
#     # 上传参数
#     data = {
#         "comment": comment,
#         "minorEdit": str(minor_edit).lower()  # Confluence API要求布尔值为小写字符串
#     }

#     # 发送上传请求
#     upload_url = f"{confluence_url}/rest/api/content/{page_id}/child/attachment"
#     try:
#         response = requests.post(upload_url, headers=headers, files=files, data=data)
#         response.raise_for_status()  # 抛出HTTP错误状态码异常
#         print("文件上传成功！")
#         return response.json()
#     except requests.exceptions.RequestException as e:
#         print(f"上传失败：{str(e)}")
#         if hasattr(e, 'response') and e.response is not None:
#             print(f"错误详情：{e.response.text}")
#         return None
#     finally:
#         # 关闭所有打开的文件句柄
#         for file_tuple in files:
#             file_tuple[1][1].close()

# # 主执行逻辑
# if __name__ == "__main__":
#     # 配置文件路径（可根据实际路径修改）
#     CONFIG_FILE = "D:\\供应链\\芯片\\updata_file.json"
#     try:
#         # 加载配置
#         config = load_confluence_config(CONFIG_FILE)
#         # 执行上传
#         result = confluence_upload_from_config(config)
#         # 打印上传结果（包含附件ID、下载链接等）
#         if result:
#             print("\n上传结果详情：")
#             for attachment in result.get("results", []):
#                 print(f"文件名：{attachment['title']}")
#                 print(f"附件ID：{attachment['id']}")
#                 print(f"下载链接：{attachment['_links']['download']}")
#                 print("-" * 50)
#     except Exception as e:
#         print(f"执行失败：{str(e)}")



import os
import json
import base64
import requests
import mimetypes

# 读取配置
def load_config():
    with open("D:\\供应链\\芯片\\updata_file.json", 'r', encoding='utf-8') as f:
        return json.load(f)

# 核心PAT上传
def upload():
    config = load_config()
    url = f"{config['confluence_url']}/rest/api/content/{config['page_id']}/child/attachment"
    # PAT基础认证：用户名+PAT拼接后Base64编码
    auth_str = f"{config['server_auth']['username']}:{config['server_auth']['password']}"
    auth = base64.b64encode(auth_str.encode()).decode()

    # 准备请求头（必加的两个核心头）
    headers = {
        "Authorization": f"Basic {auth}",
        "X-Atlassian-Token": "no-check"
    }

    # 准备文件（单文件/多文件均可）
    files = []
    for fp in config['upload_config']['file_paths']:
        if not os.path.exists(fp):
            print(f"文件不存在：{fp}")
            return
        fname = os.path.basename(fp)
        mtype = mimetypes.guess_type(fp)[0] or "application/octet-stream"
        files.append(("file", (fname, open(fp, "rb"), mtype)))

    # 上传参数
    data = {
        "comment": config['upload_config']['comment'],
        "minorEdit": str(config['upload_config']['minor_edit']).lower()
    }

    # 发送请求（跳过内网证书验证，关键）
    try:
        res = requests.post(
            url,
            headers=headers,
            files=files,
            data=data,
            verify=False  # 企业内网自签证书必加
        )
        res.raise_for_status()
        print("✅ 上传成功！")
        print("📌 附件信息：", res.json())
        # 关闭文件
        for f in files:
            f[1][1].close()
    except Exception as e:
        print(f"❌ 上传失败：{str(e)}")
        if hasattr(e, 'response') and e.response:
            print(f"📝 错误详情：{e.response.text}")

if __name__ == "__main__":
    upload()
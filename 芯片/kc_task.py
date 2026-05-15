import json
import os
import subprocess
import sys
from datetime import datetime,timedelta

# ===================== 配置项（请根据你的实际情况修改）=====================
CONFIG_FILE_PATH = r"D:\供应链\芯片\configkcV8.json"  # 你的配置文件路径
MAIN_PY_FILE_PATH = r"D:\供应链\芯片\V8.py"  # 你的主Python文件路径
DATE_FIELD_PATH = "email_config.search_criteria"  # 配置文件中要修改的时间字段路径（按需调整）
PYTHON_EXECUTABLE = sys.executable  # Python解释器路径（默认用当前环境的Python）
# ===========================================================================

def update_config_date(config_path: str, date_field_path: str) -> bool:
    """
    修改配置文件中的时间字段为当天日期（IMAP搜索格式：DD-Mmm-YYYY，如 12-Feb-2026）
    :param config_path: 配置文件路径
    :param date_field_path: 要修改的字段路径，如 "email_config.search_criteria"
    :return: 是否修改成功
    """
    try:
        # 1. 读取配置文件
        with open(config_path, "r", encoding="utf-8") as f:
            config = json.load(f)
        
        # 2. 解析字段路径，定位到要修改的字段（支持多级路径，如 a.b.c）
        field_parts = date_field_path.split(".")
        current_node = config
        for part in field_parts[:-1]:
            if part not in current_node:
                raise ValueError(f"配置文件中未找到字段：{part}")
            current_node = current_node[part]
        
        # 3. 生成当天的IMAP格式日期（如 12-Feb-2026）
        today = datetime.now()
        month_abbr = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", 
                      "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"][today.month-1]
        imap_date_str = f'SINCE "{today.day:02d}-{month_abbr}-{today.year}" BEFORE "{(today + timedelta(days=1)).day:02d}-{month_abbr}-{today.year}"'
        
        # 4. 修改字段值
        current_node[field_parts[-1]] = imap_date_str
        print(f"✅ 已将配置文件时间修改为：{imap_date_str}")
        
        # 5. 保存配置文件（保留原有格式和注释）
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        
        return True
    except Exception as e:
        print(f"❌ 修改配置文件失败：{str(e)}")
        return False

def run_main_python_file(py_file_path: str) -> bool:
    """执行主Python文件（彻底解决编码/类型错误）"""
    try:
        if not os.path.exists(py_file_path):
            raise FileNotFoundError(f"主Python文件不存在：{py_file_path}")
        
        # 执行主程序（用当前Python环境）
        result = subprocess.run(
            [PYTHON_EXECUTABLE, py_file_path],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            # 关键：不指定encoding，强制返回字节流；同时忽略执行中的编码错误
            encoding=None,  
            errors="ignore",
            timeout=3600  # 超时时间1小时，可根据你的程序调整
        )
        
        # 安全处理输出：先判断类型，bytes才解码，str直接用
        def safe_process_output(output):
            if output is None:
                return ""
            # 如果是字节流，尝试解码；如果是字符串，直接返回
            if isinstance(output, bytes):
                try:
                    return output.decode("gbk", errors="ignore")  # Windows优先GBK
                except:
                    return output.decode("utf-8", errors="ignore")
            elif isinstance(output, str):
                return output
            else:
                return str(output, errors="ignore")
        
        stdout_str = safe_process_output(result.stdout)
        stderr_str = safe_process_output(result.stderr)
        
        if result.returncode == 0:
            print(f"✅ 主程序执行成功：\n{stdout_str}")
            return True
        else:
            print(f"❌ 主程序执行失败，错误信息：\n{stderr_str}")
            return False
    except Exception as e:
        print(f"❌ 执行主程序异常：{str(e)}")
        # 打印异常详情，方便排查
        import traceback
        print(f"📝 异常详情：\n{traceback.format_exc()}")
        return False
if __name__ == "__main__":
    print(f"========== 定时任务开始执行（{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}）==========")
    
    # 第一步：修改配置文件时间
    if not update_config_date(CONFIG_FILE_PATH, DATE_FIELD_PATH):
        sys.exit(1)
    
    # 第二步：执行主程序
    if not run_main_python_file(MAIN_PY_FILE_PATH):
        sys.exit(1)
    
    print(f"========== 定时任务执行完成（{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}）==========")
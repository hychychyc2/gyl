from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import WebDriverException, TimeoutException
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from time import sleep

# -------------------------- 仅需修改这2个参数 --------------------------
TARGET_URL = "https://icop.bitmain.vip/home"  # 目标内网网页（无需改）
BROWSER_TYPE = "Edge"  # 保持Edge，无需换Chrome

# -------------------------- 验证脚本核心逻辑 --------------------------
def verify_selenium_network():
    print("="*50)
    print("📡 Selenium内网网络环境验证脚本")
    print("="*50)
    print(f"目标验证网页：{TARGET_URL}")
    print("验证逻辑：Selenium启动Edge，继承系统代理/VPN，测试能否访问内网")
    print("="*50)

    # 1. 初始化浏览器（带系统代理配置，关键！）
    try:
        print("\n1. 正在启动浏览器（启用系统代理/VPN继承）...")
        options = webdriver.EdgeOptions()
        options.add_argument("--proxy-server=system")  # 核心：继承系统代理
        options.add_argument("--no-sandbox")  # 内网环境权限兼容
        options.add_argument("--disable-dev-shm-usage")  # 避免内存问题
        options.add_experimental_option("detach", True)  # 不自动关闭浏览器，方便手动查看
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        
        service = Service(EdgeChromiumDriverManager().install())
        driver = webdriver.Edge(service=service, options=options)
        driver.maximize_window()
        print("✅ 浏览器启动成功")
    except Exception as e:
        print(f"❌ 浏览器启动失败：{str(e)}")
        print("排查建议：1. 检查Edge浏览器是否正常安装 2. 关闭占用Edge的进程 3. 重新安装webdriver-manager")
        return

    # 2. 尝试访问目标内网网页
    try:
        print(f"\n2. 正在访问内网网页：{TARGET_URL}...")
        driver.get(TARGET_URL)
        
        # 等待页面加载，验证是否成功（检测标题非空或页面有内容）
        try:
            WebDriverWait(driver, 20).until(lambda d: d.title != "")
            page_title = driver.title
            print(f"✅ 网页访问成功！页面标题：{page_title}")
            print("\n🎉 网络环境验证通过！")
            print("结论：Selenium浏览器已成功继承系统代理/VPN，可正常访问内网")
            print("后续操作：直接运行原业务脚本即可，无需额外配置")
        except TimeoutException:
            print(f"❌ 网页加载超时（20秒），未检测到页面内容")
            print("\n排查建议：")
            print("1. 手动在Selenium打开的浏览器中输入目标URL，看是否能加载")
            print("2. 若手动也不能加载：检查VPN是否连接、是否已登录，联系IT放行Edge进程")
            print("3. 若手动能加载：延长原脚本的WAIT_TIME（建议30秒），或检查URL是否正确")
    except WebDriverException as e:
        print(f"❌ 访问网页失败：{str(e)}")
        print("\n排查建议：")
        print("1. 确认VPN已连接，且未设置“仅允许手动程序访问”")
        print("2. 关闭浏览器代理设置（Selenium已启用系统代理，无需额外配置）")
        print("3. 检查防火墙是否拦截Selenium/Edge进程，添加例外")

    # 3. 结束提示
    print("\n" + "="*50)
    print("📌 操作提示：")
    print("1. 不要关闭浏览器，可手动测试其他内网链接")
    print("2. 验证完成后，手动关闭浏览器即可")
    print("="*50)

if __name__ == "__main__":
    verify_selenium_network()
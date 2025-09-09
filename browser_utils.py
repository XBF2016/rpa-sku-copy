import os
import time
import shutil
import subprocess
from pathlib import Path
from typing import Optional
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.edge.service import Service as EdgeService

from common import read_browser_path, log_debug
from sku_utils import SKU_ITEM_SELECTOR

# 浏览器与驱动管理工具


def is_process_running(image_name: str) -> bool:
    """使用 Windows 'tasklist' 命令检查给定进程名是否在运行。"""
    try:
        result = subprocess.run(
            ["tasklist", "/FI", f"IMAGENAME eq {image_name}"],
            capture_output=True,
            text=True,
            check=False,
        )
        stdout = (result.stdout or "").lower()
        return image_name.lower() in stdout
    except Exception:
        return False


def kill_edge_processes() -> None:
    """强制结束 Edge 相关进程（msedge.exe、msedgewebview2.exe）。"""
    for image in ("msedge.exe", "msedgewebview2.exe"):
        try:
            subprocess.run(["taskkill", "/IM", image, "/F"], capture_output=True, text=True)
        except Exception:
            pass


def kill_driver_processes() -> None:
    """强制结束 EdgeDriver 相关进程（msedgedriver.exe）。"""
    for image in ("msedgedriver.exe",):
        try:
            subprocess.run(["taskkill", "/IM", image, "/F"], capture_output=True, text=True)
        except Exception:
            pass


def find_msedgedriver_path() -> Optional[str]:
    """查找 msedgedriver.exe 的本地路径（避免联网下载）。
    查找顺序：
      1) 环境变量 MSEDGEDRIVER 指定的路径
      2) 项目根目录下的 driver/msedgedriver.exe
      3) 与 browser.txt 中的 msedge.exe 同目录下的 msedgedriver.exe
      4) 常见安装目录
      5) 系统 PATH 中的 msedgedriver
    """
    # 1) 环境变量
    env_path = os.environ.get("MSEDGEDRIVER", "").strip('"')
    if env_path:
        if os.path.isfile(env_path):
            return env_path
        cand = Path(env_path) / "msedgedriver.exe"
        if cand.exists():
            return str(cand)

    # 2) 项目 driver 目录
    cand = Path(__file__).resolve().parent / "driver" / "msedgedriver.exe"
    if cand.exists():
        return str(cand)

    # 3) 与 msedge.exe 同目录
    edge_exe = read_browser_path()
    if edge_exe and os.path.exists(edge_exe):
        cand = Path(edge_exe).parent / "msedgedriver.exe"
        if cand.exists():
            return str(cand)

    # 4) 常见安装目录
    common_paths = [
        r"C:\\Program Files\\Microsoft\\Edge\\Application\\msedgedriver.exe",
        r"C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedgedriver.exe",
    ]
    for p in common_paths:
        if os.path.exists(p):
            return p

    # 5) PATH
    which = shutil.which("msedgedriver")
    if which:
        return which

    return None


def prepare_clean_edge_state() -> None:
    """准备干净的 Edge 运行环境，确保不受残留进程影响。"""
    print("[步骤] 检查并关闭现有的 Edge 进程...")
    t_all0 = time.perf_counter()
    if is_process_running("msedge.exe"):
        print("[信息] 检测到 Edge 正在运行，关闭所有 Edge 相关进程...")
        t0 = time.perf_counter()
        kill_edge_processes()
        t0 = time.perf_counter()
        while time.perf_counter() - t0 < 1.0:
            if not is_process_running("msedge.exe"):
                break
            time.sleep(0.1)
        log_debug(f"关闭 Edge 进程总耗时 {(time.perf_counter() - t0)*1000:.0f}ms")
    print("[步骤] 关闭可能残留的 EdgeDriver 进程...")
    t1 = time.perf_counter()
    kill_driver_processes()
    log_debug(f"关闭 EdgeDriver 进程耗时 {(time.perf_counter() - t1)*1000:.0f}ms；清理总耗时 {(time.perf_counter() - t_all0):.3f}s")


def init_edge_driver() -> webdriver.Edge:
    """初始化 Edge WebDriver（使用本地 driver 与用户登录态）。"""
    user_data_dir = os.path.expanduser("~\\AppData\\Local\\Microsoft\\Edge\\User Data")

    print("[步骤] 初始化 Edge WebDriver（使用用户登录态）...")
    t0 = time.perf_counter()
    options = EdgeOptions()
    options.add_argument("--start-maximized")
    options.add_argument(f"--user-data-dir={user_data_dir}")
    options.add_argument("--profile-directory=Default")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])  # 排除日志开关
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument("--log-level=3")
    options.add_argument("--disable-logging")
    options.add_argument("--disable-quic")
    options.add_argument("--ignore-certificate-errors")
    options.set_capability("acceptInsecureCerts", True)
    try:
        options.set_capability("pageLoadStrategy", "eager")
    except Exception:
        pass

    driver_path = find_msedgedriver_path()
    if not driver_path:
        print("[错误] 未找到 msedgedriver.exe")
        raise RuntimeError("未找到本地 EdgeDriver")

    print(f"[信息] 使用本地 EdgeDriver: {driver_path}")
    try:
        service = EdgeService(executable_path=driver_path, log_output=subprocess.DEVNULL)
    except TypeError:
        service = EdgeService(executable_path=driver_path)

    t1 = time.perf_counter()
    driver = webdriver.Edge(service=service, options=options)
    driver.implicitly_wait(1)
    try:
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    except Exception:
        pass
    log_debug(f"初始化 EdgeDriver 耗时 {(time.perf_counter() - t0):.3f}s（创建Service {(t1 - t0)*1000:.0f}ms，启动Driver {(time.perf_counter() - t1)*1000:.0f}ms）")
    return driver


def open_product_page(driver: webdriver.Edge, url: str, wait_timeout: int = 30) -> None:
    """打开商品页并等待 SKU 区域出现。"""

    print("[步骤] 打开商品页面...")
    t0 = time.perf_counter()
    driver.get(url)
    t_get = time.perf_counter()
    time.sleep(0.3)
    t_sleep = time.perf_counter()
    WebDriverWait(driver, wait_timeout).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, SKU_ITEM_SELECTOR))
    )
    t_wait = time.perf_counter()
    log_debug(
        f"打开商品页: get(url) {(t_get - t0)*1000:.0f}ms；随机等待 {(t_sleep - t_get)*1000:.0f}ms；等待SKU出现 {(t_wait - t_sleep)*1000:.0f}ms；总计 {(t_wait - t0):.3f}s"
    )

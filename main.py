# -*- coding: utf-8 -*-
"""
抖店-创建规格维度 RPA

说明：
- 从 conf/douyin/product-url.txt 读取商品草稿链接（若不存在则回退 conf/product-url.txt）
- 从 conf/规格.yml 读取需要创建的“维度”（仅顶层键名，不创建选项）
- 打开浏览器访问草稿页，依次点击“添加规格类型”，展开“规格类型下拉按钮”，在下拉列表中选择维度；
  若列表没有该维度，则点击“创建类型”，在弹出的输入框中输入维度名并回车。
- 代码内全部中文日志与注释，严格参考 “元素示例/” 下的文件来定位元素。

注意：
- 本实现不依赖 PyYAML，为减少依赖与改动，采用最简行解析从 conf/规格.yml 提取顶层键作为维度名。
- 若页面需登录，请先手动完成登录后再继续。
"""
from __future__ import annotations

import os
import sys
import time
from pathlib import Path
import subprocess
import socket
from typing import List, Optional, Set

from robocorp.tasks import task
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.common.exceptions import SessionNotCreatedException, WebDriverException
from selenium.webdriver.edge.options import Options as EdgeOptions

# -------------------------
# 路径与常量
# -------------------------
PROJECT_ROOT = Path(__file__).resolve().parent
CONF_DIR = PROJECT_ROOT / "conf"
DRIVER_PATH = PROJECT_ROOT / "driver" / "msedgedriver.exe"
BROWSER_PATH_FILE = CONF_DIR / "browser.txt"

PRODUCT_URL_FILE_PRI = CONF_DIR / "douyin" / "product-url.txt"
PRODUCT_URL_FILE_FALLBACK = CONF_DIR / "product-url.txt"
SPECS_YAML_FILE = CONF_DIR / "规格.yml"

# 元素选择器（严格参考 元素示例/ 下的文件）
# - 添加规格类型 按钮（元素示例/添加规格类型.html）
X_ADD_SPEC_BUTTON = (
    By.XPATH,
    "//button[contains(@class,'ecom-g-btn-dashed')]//span[contains(normalize-space(.),'添加规格类型')]/ancestor::button",
)
# - 规格类型下拉按钮（元素示例/规格类型下拉按钮.html）: 占位文案“请选择规格类型”
X_SPEC_DROPDOWN_TRIGGER = (
    By.XPATH,
    "//div[contains(@class,'ecom-g-select') and contains(@class,'ecom-g-select-single') and .//span[contains(@class,'ecom-g-select-selection-placeholder') and contains(.,'请选择规格类型')]]",
)
# - 规格类型下拉列表（元素示例/规格类型下拉列表.html）: 可见的下拉菜单容器
X_VISIBLE_DROPDOWN = (
    By.XPATH,
    "//div[contains(@class,'ecom-g-select-dropdown') and not(contains(@class,'ecom-g-select-dropdown-hidden'))]",
)
# - 下拉列表项内容容器 class: ecom-g-select-item-option-content
X_DROPDOWN_ITEM_BY_TEXT_TMPL = (
    By.XPATH,
    ".//div[contains(@class,'ecom-g-select-item-option-content') and normalize-space(text())='{text}']",
)
# - 下拉列表底部“创建类型”链接
X_CREATE_TYPE_LINK = (By.XPATH, ".//a[normalize-space(.)='创建类型']")
# - “请输入规格类型” 输入框（描述中给出）
X_INPUT_CREATE_TYPE = (
    By.XPATH,
    "//input[contains(@placeholder,'请输入规格类型')]",
)
# - 页面中已添加的维度展示（参考 元素示例/商品规格区域.html 中的 ecom-g-select-selection-item）
X_EXISTING_SPEC_ITEMS = (
    By.XPATH,
    "//span[contains(@class,'ecom-g-select-selection-item')]",
)


def read_browser_path() -> Optional[str]:
    """读取浏览器可执行文件路径（Edge）。
    优先读取 conf/browser.txt；若不存在则返回 None（由 EdgeDriver 自行查找）。
    """
    try:
        if BROWSER_PATH_FILE.exists():
            p = BROWSER_PATH_FILE.read_text(encoding="utf-8").strip()
            if p:
                print(f"[信息] 已读取浏览器路径: {p}")
                return p
    except Exception as e:
        print(f"[警告] 读取浏览器路径失败: {e}")
    return None


def read_product_url() -> str:
    """读取商品草稿页链接。
    优先 conf/douyin/product-url.txt，回退 conf/product-url.txt。
    """
    for fp in (PRODUCT_URL_FILE_PRI, PRODUCT_URL_FILE_FALLBACK):
        if fp.exists():
            url = fp.read_text(encoding="utf-8").strip()
            if url:
                print(f"[信息] 使用商品链接: {fp} -> {url}")
                return url
    raise RuntimeError("未找到商品链接，请在 conf/douyin/product-url.txt 或 conf/product-url.txt 填写链接")


def read_spec_dimensions_from_yaml(yaml_path: Path) -> List[str]:
    """最简 YAML 顶层键解析：提取作为“维度”的键名列表。
    - 不依赖 PyYAML；仅处理类似如下结构：
        颜色分类:
          - A
          - B
        尺码:
          - M
          - L
    - 只取顶层键名 [颜色分类, 尺码]
    """
    if not yaml_path.exists():
        raise RuntimeError(f"未找到规格配置文件: {yaml_path}")
    dims: List[str] = []
    try:
        with yaml_path.open("r", encoding="utf-8") as f:
            for raw in f:
                line = raw.rstrip("\n\r")
                if not line.strip():
                    continue
                if line.lstrip().startswith("#"):
                    continue
                # 仅识别顶层（不以空格开头），并以冒号结尾
                if not line.startswith(" ") and line.strip().endswith(":"):
                    key = line.strip()[:-1].strip()
                    if key:
                        dims.append(key)
    except Exception as e:
        raise RuntimeError(f"解析规格配置失败，请检查 conf/规格.yml 格式是否为顶层键: 列表 的结构。错误: {e}")
    if not dims:
        raise RuntimeError("未从 conf/规格.yml 解析到任何维度，请确认顶层键是否存在")
    print(f"[信息] 将要创建的规格维度: {dims}")
    return dims


def build_edge_driver() -> webdriver.Edge:
    """构建 Edge WebDriver。优先使用本地 driver/msedgedriver.exe 与 conf/browser.txt。"""
    opts = EdgeOptions()
    opts.add_argument("start-maximized")
    # 部分稳定启动参数，减少首次启动与提示
    opts.add_argument("--no-first-run")
    opts.add_argument("--no-default-browser-check")
    opts.add_argument("--disable-popup-blocking")
    opts.add_argument("--remote-allow-origins=*")
    # 复用本机 Edge 登录态（优先环境变量，其次本机默认路径）
    user_data_dir = os.environ.get("EDGE_USER_DATA_DIR")
    profile_dir = os.environ.get("EDGE_PROFILE", "Default")
    if not user_data_dir:
        lad = os.environ.get("LOCALAPPDATA")
        if lad:
            default_ud = Path(lad) / "Microsoft" / "Edge" / "User Data"
            if default_ud.exists():
                user_data_dir = str(default_ud)
    if user_data_dir:
        opts.add_argument(f"--user-data-dir={user_data_dir}")
        if profile_dir:
            opts.add_argument(f"--profile-directory={profile_dir}")
        print(f"[信息] 将复用 Edge 登录态：{user_data_dir} / {profile_dir}")
        print("[提示] 如 Edge 已经打开，可能因同一用户数据目录被占用而导致启动失败，请先关闭所有 Edge 窗口再运行；或改用“附加模式”（见 README）。")
    else:
        print("[提示] 未找到 Edge 用户数据目录，将以临时会话启动（可能需要登录）")
    browser_path = read_browser_path()
    # 不强制指定 binary_location，避免与已安装 Edge 版本不匹配引发崩溃（DevToolsActivePort 不存在）

    # 附加模式：如果提供了 EDGE_ATTACH_DEBUG_ADDR，则附加到已启动的 Edge（需用户以该端口启动 Edge）
    attach_addr = os.environ.get("EDGE_ATTACH_DEBUG_ADDR")
    if attach_addr:
        try:
            # 附加模式不需要 service
            opts.add_experimental_option("debuggerAddress", attach_addr)
            print(f"[信息] 正在以附加模式连接到 Edge：{attach_addr}")
            driver = webdriver.Edge(options=opts)
            driver.implicitly_wait(2)
            return driver
        except Exception as e:
            print(f"[错误] 附加模式连接失败：{e}")
            print("[提示] 请确认已用 --remote-debugging-port 启动了 Edge，并确保端口可用")

    service = EdgeService(executable_path=str(DRIVER_PATH)) if DRIVER_PATH.exists() else None
    def _try_attach_mode_on_failure(e: Exception):
        msg = str(e)
        if "DevToolsActivePort" in msg or "failed to start" in msg or "crashed" in msg:
            print("[警告] Edge 启动失败（DevToolsActivePort/failed to start/crashed）。尝试以“远程调试附加模式”重新连接…")
            # 避免用户数据目录被占用：若已有 Edge 运行，请提示关闭
            if is_edge_running():
                print("[提示] 检测到系统中已有 Edge 正在运行，为避免用户数据目录占用，请关闭所有 Edge 窗口后按回车继续…")
                try:
                    input()
                except Exception:
                    time.sleep(3)
            # 自动启动一个带远程调试端口的 Edge，并复用用户数据目录
            port = get_free_port()
            try:
                start_edge_debug_and_wait(browser_path, user_data_dir, profile_dir, port, timeout=20)
                # 以附加模式连接
                opts.add_experimental_option("debuggerAddress", f"127.0.0.1:{port}")
                drv = webdriver.Edge(options=opts)
                drv.implicitly_wait(2)
                return drv
            except Exception as e2:
                print(f"[错误] 自动启动并附加 Edge 失败：{e2}")
                raise
        else:
            raise e

    if service is None:
        print("[警告] 未找到 driver/msedgedriver.exe，将尝试自动匹配驱动（Selenium Manager）")
        try:
            driver = webdriver.Edge(options=opts)
        except (SessionNotCreatedException, WebDriverException) as e:
            driver = _try_attach_mode_on_failure(e)
    else:
        # 先尝试使用本地驱动，失败再自动回退
        try:
            driver = webdriver.Edge(service=service, options=opts)
        except (SessionNotCreatedException, WebDriverException) as e:
            msg = str(e)
            if "DevToolsActivePort" in msg or "failed to start" in msg or "crashed" in msg:
                print("[警告] 使用本地 EdgeDriver 启动失败，可能与浏览器版本不匹配，将回退为自动匹配驱动…")
                try:
                    driver = webdriver.Edge(options=opts)
                except (SessionNotCreatedException, WebDriverException) as e2:
                    driver = _try_attach_mode_on_failure(e2)
            else:
                raise
    driver.implicitly_wait(2)
    return driver


def get_free_port() -> int:
    s = socket.socket()
    s.bind(("127.0.0.1", 0))
    port = s.getsockname()[1]
    s.close()
    return port


def is_edge_running() -> bool:
    try:
        # Windows: 查询是否有 msedge.exe 进程
        out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq msedge.exe"], capture_output=True, text=True)
        return "msedge.exe" in (out.stdout or "")
    except Exception:
        return False


def start_edge_debug_and_wait(msedge_path: Optional[str], user_data_dir: Optional[str], profile_dir: str, port: int, timeout: int = 15) -> None:
    exe = msedge_path or "msedge.exe"
    args = [
        exe,
        f"--remote-debugging-port={port}",
    ]
    if user_data_dir:
        args.append(f"--user-data-dir={user_data_dir}")
    if profile_dir:
        args.append(f"--profile-directory={profile_dir}")
    # 启动 Edge（不阻塞）
    print(f"[信息] 启动 Edge 远程调试实例：{exe}，端口 {port}，用户数据 {user_data_dir} / {profile_dir}")
    subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    # 等端口就绪
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            with socket.create_connection(("127.0.0.1", port), timeout=1):
                print("[信息] 远程调试端口已就绪")
                return
        except Exception:
            time.sleep(0.3)
    raise RuntimeError("远程调试端口未在预期时间内就绪，请检查 Edge 是否成功启动")


def wait_for_login_and_page_ready(driver: webdriver.Edge, timeout: int = 60, max_total_seconds: int = 600) -> None:
    """等待页面出现“添加规格类型”按钮。
    - 若登录过期/未登录，请在浏览器完成登录，程序会循环等待，直到检测到目标元素或超时（默认10分钟）。
    """
    deadline = time.time() + max_total_seconds
    has_prompted = False
    while time.time() < deadline:
        try:
            WebDriverWait(driver, timeout).until(EC.presence_of_element_located(X_ADD_SPEC_BUTTON))
            print("[信息] 页面已加载，检测到“添加规格类型”按钮")
            return
        except Exception:
            if not has_prompted:
                print("[提示] 检测到登录状态可能失效或页面尚未就绪，请在浏览器完成登录；程序会持续等待，直到检测到“添加规格类型”按钮或超时（最多10分钟）…")
                has_prompted = True
            time.sleep(2)
    raise RuntimeError("长时间未检测到“添加规格类型”按钮，可能未完成登录或页面异常")


def navigate_to_url(driver: webdriver.Edge, url: str, wait_seconds: int = 5) -> None:
    """稳健导航到指定 URL：
    1) 先使用 driver.get(url)
    2) 若短暂等待后仍停留在 about:blank/空白，尝试用 window.open 打开新标签并切换
    3) 仍不行则用脚本设置 location.href
    """
    print(f"[步骤] 导航到链接（driver.get）：{url}")
    try:
        driver.get(url)
    except Exception as e:
        print(f"[警告] driver.get 抛出异常：{e}，将尝试使用备用方案")
    time.sleep(1)
    try:
        cur = (driver.current_url or "").strip()
    except Exception:
        cur = ""
    if not cur or cur == "about:blank":
        print("[提示] 当前页面为空白，尝试使用 window.open 打开新标签…")
        try:
            driver.execute_script("window.open(arguments[0], '_blank');", url)
            time.sleep(0.5)
            handles = driver.window_handles
            driver.switch_to.window(handles[-1])
            time.sleep(0.5)
        except Exception as e:
            print(f"[警告] window.open 失败：{e}，改用 location.href …")
            try:
                driver.execute_script("location.href = arguments[0];", url)
            except Exception as e2:
                print(f"[错误] location.href 也失败：{e2}")

def collect_existing_dimensions(driver: webdriver.Edge) -> Set[str]:
    """收集页面中已存在的维度名（从已选择项的展示中提取文本）。"""
    existed: Set[str] = set()
    try:
        nodes = driver.find_elements(*X_EXISTING_SPEC_ITEMS)
        for n in nodes:
            txt = (n.text or n.get_attribute("title") or "").strip()
            if txt:
                existed.add(txt)
    except Exception:
        pass
    print(f"[信息] 当前已存在的维度: {list(existed)}")
    return existed


def click_add_spec_button(driver: webdriver.Edge) -> None:
    btn = WebDriverWait(driver, 20).until(EC.element_to_be_clickable(X_ADD_SPEC_BUTTON))
    btn.click()
    print("[步骤] 已点击“添加规格类型”按钮")


def open_latest_spec_dropdown(driver: webdriver.Edge) -> None:
    """点击“请选择规格类型”的下拉触发器（若有多个，点击最后一个）。"""
    nodes = WebDriverWait(driver, 20).until(EC.presence_of_all_elements_located(X_SPEC_DROPDOWN_TRIGGER))
    nodes[-1].click()
    print("[步骤] 已点击“规格类型下拉按钮”（占位：请选择规格类型）")


def select_dimension_from_dropdown(driver: webdriver.Edge, dim: str) -> bool:
    """在下拉列表中选择指定维度；若不存在则返回 False。"""
    dropdown = WebDriverWait(driver, 20).until(EC.presence_of_element_located(X_VISIBLE_DROPDOWN))
    # 精确匹配项
    by, tmpl = X_DROPDOWN_ITEM_BY_TEXT_TMPL
    locator = (by, tmpl.format(text=dim))
    try:
        item = WebDriverWait(dropdown, 2).until(lambda d: dropdown.find_element(*locator))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", item)
        item.click()
        print(f"[步骤] 已在下拉列表中选择维度：{dim}")
        return True
    except Exception:
        print(f"[信息] 下拉列表中未找到维度：{dim}，准备走“创建类型”流程…")
        return False


def create_dimension_via_dialog(driver: webdriver.Edge, dim: str) -> None:
    """点击“创建类型”，在输入框中输入维度名并回车。"""
    dropdown = WebDriverWait(driver, 10).until(EC.presence_of_element_located(X_VISIBLE_DROPDOWN))
    # 点击“创建类型”
    try:
        create_link = dropdown.find_element(*X_CREATE_TYPE_LINK)
        create_link.click()
        print("[步骤] 已点击“创建类型”")
    except Exception:
        print("[错误] 未在下拉中找到“创建类型”链接，无法创建新维度")
        raise
    # 输入框输入维度
    try:
        input_box = WebDriverWait(driver, 10).until(EC.presence_of_element_located(X_INPUT_CREATE_TYPE))
        input_box.clear()
        input_box.send_keys(dim)
        input_box.send_keys(Keys.ENTER)
        print(f"[步骤] 已输入新维度并回车：{dim}")
    except Exception:
        print("[错误] 未找到“请输入规格类型”的输入框，创建失败")
        raise
    # 等待选择结果反映到页面（任一位置出现该维度的已选择项）
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, f"//span[contains(@class,'ecom-g-select-selection-item') and normalize-space(text())='{dim}']"))
    )


@task
def create_douyin_spec_dimensions() -> None:
    """创建抖店商品规格维度（仅维度，不创建选项）。"""
    print("[开始] 抖店-创建规格维度 任务启动…")
    url = read_product_url()
    dims = read_spec_dimensions_from_yaml(SPECS_YAML_FILE)

    driver = build_edge_driver()
    try:
        print(f"[步骤] 打开商品草稿链接: {url}")
        navigate_to_url(driver, url)
        wait_for_login_and_page_ready(driver, timeout=60)

        existed = collect_existing_dimensions(driver)

        for dim in dims:
            if dim in existed:
                print(f"[跳过] 维度已存在：{dim}")
                continue
            # 依次添加
            click_add_spec_button(driver)
            open_latest_spec_dropdown(driver)
            if not select_dimension_from_dropdown(driver, dim):
                create_dimension_via_dialog(driver, dim)
            # 更新已存在集合
            time.sleep(0.5)
            existed = collect_existing_dimensions(driver)

        print("[完成] 规格维度创建流程结束。")
    finally:
        # 为便于人工检查，保留窗口 3 秒，可视需要调整
        time.sleep(3)
        try:
            driver.quit()
        except Exception:
            pass


if __name__ == "__main__":
    # 允许直接运行：python main.py 将执行该任务
    create_douyin_spec_dimensions()

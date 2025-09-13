# -*- coding: utf-8 -*-
"""
浏览器启动与页面通用流程工具：

- 构建 Edge WebDriver（含多层回退：本地驱动 -> 自动匹配 -> 远程调试附加 -> 临时用户数据目录）
- 稳健导航到目标 URL
- 等待页面就绪（等待“添加规格类型”按钮出现）

严格使用中文日志与报错；遵循项目既有的容错与回退逻辑。
"""
from __future__ import annotations

import os
import time
import socket
import subprocess
from typing import Optional

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.common.exceptions import SessionNotCreatedException, WebDriverException
from selenium.webdriver.edge.options import Options as EdgeOptions

from config import (
    DRIVER_PATH,
    TEMP_DIR,
    read_browser_path,
)


# 限制：尽量优先复用已有工具；若导入失败则本文件内提供兜底实现
try:
    from browser_utils import kill_edge_processes
except Exception:
    def kill_edge_processes() -> None:
        """强制结束 Edge 相关进程（msedge.exe、msedgewebview2.exe）。"""
        try:
            for image in ("msedge.exe", "msedgewebview2.exe"):
                try:
                    subprocess.run(["taskkill", "/IM", image, "/F"], capture_output=True, text=True)
                except Exception:
                    pass
        except Exception:
            pass


def _make_edge_options(user_data_dir: Optional[str], profile_dir: Optional[str], debugger_addr: Optional[str] = None) -> EdgeOptions:
    opts = EdgeOptions()
    opts.add_argument("start-maximized")
    opts.add_argument("--no-first-run")
    opts.add_argument("--no-default-browser-check")
    opts.add_argument("--disable-popup-blocking")
    opts.add_argument("--remote-allow-origins=*")
    if user_data_dir:
        opts.add_argument(f"--user-data-dir={user_data_dir}")
        if profile_dir:
            opts.add_argument(f"--profile-directory={profile_dir}")
    if debugger_addr:
        opts.add_experimental_option("debuggerAddress", debugger_addr)
    return opts


def build_edge_driver() -> webdriver.Edge:
    """构建 Edge WebDriver。优先使用本地 driver/msedgedriver.exe 与 conf/browser.txt。"""
    # 先关闭所有已运行的 Edge，避免用户数据目录被占用导致启动崩溃
    try:
        print("[步骤] 正在关闭已运行的 Edge 进程…")
        kill_edge_processes()
        time.sleep(0.5)
    except Exception:
        pass

    # 偏好复用用户数据目录（复用本机 Edge 登录态）
    user_data_dir = os.environ.get("EDGE_USER_DATA_DIR")
    profile_dir = os.environ.get("EDGE_PROFILE", "Default")
    if not user_data_dir:
        lad = os.environ.get("LOCALAPPDATA")
        if lad:
            from pathlib import Path
            default_ud = Path(lad) / "Microsoft" / "Edge" / "User Data"
            if default_ud.exists():
                user_data_dir = str(default_ud)
    opts = _make_edge_options(user_data_dir, profile_dir)
    if user_data_dir:
        print(f"[信息] 将复用 Edge 登录态：{user_data_dir} / {profile_dir}")
        print("[提示] 如 Edge 已经打开，可能因同一用户数据目录被占用而导致启动失败，请先关闭所有 Edge 窗口再运行；或改用“附加模式”（见 README）。")
    else:
        print("[提示] 未找到 Edge 用户数据目录，将以临时会话启动（可能需要登录）")
    browser_path = read_browser_path()

    # 附加模式：如果提供了 EDGE_ATTACH_DEBUG_ADDR，则附加到已启动的 Edge（需用户以该端口启动 Edge）
    attach_addr = os.environ.get("EDGE_ATTACH_DEBUG_ADDR")
    if attach_addr:
        try:
            opts = _make_edge_options(None, None, attach_addr)
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
        # 情况一：用户数据目录被占用
        if ("already in use" in msg) or ("user data directory" in msg and "in use" in msg) or ("--user-data-dir" in msg and "in use" in msg):
            try:
                TEMP_DIR.mkdir(parents=True, exist_ok=True)
            except Exception:
                pass
            # 使用时间戳创建唯一的临时用户数据目录，避免并发冲突
            tmp_ud = str(TEMP_DIR / f"edge-user-data-{int(time.time())}")
            print(f"[警告] 检测到用户数据目录被占用，将回退为临时用户数据目录启动：{tmp_ud}")
            try:
                opts_tmp = _make_edge_options(tmp_ud, "Default")
                drv = webdriver.Edge(options=opts_tmp)
                drv.implicitly_wait(2)
                return drv
            except Exception as e_tmp:
                print(f"[错误] 使用临时用户数据目录启动失败：{e_tmp}")
                # 继续走附加模式回退

        # 情况二：DevToolsActivePort/failed to start/crashed，尝试远程调试附加模式
        if "DevToolsActivePort" in msg or "failed to start" in msg or "crashed" in msg:
            print("[警告] Edge 启动失败（DevToolsActivePort/failed to start/crashed）。尝试以“远程调试附加模式”重新连接…")
            if is_edge_running():
                print("[提示] 检测到系统中已有 Edge 正在运行，为避免用户数据目录占用，请关闭所有 Edge 窗口后按回车继续…")
                try:
                    input()
                except Exception:
                    time.sleep(3)
            port = get_free_port()
            try:
                start_edge_debug_and_wait(browser_path, user_data_dir, profile_dir, port, timeout=20)
                opts2 = _make_edge_options(None, None, f"127.0.0.1:{port}")
                drv = webdriver.Edge(options=opts2)
                drv.implicitly_wait(2)
                return drv
            except Exception as e2:
                print(f"[错误] 自动启动并附加 Edge 失败：{e2}")
                try:
                    TEMP_DIR.mkdir(parents=True, exist_ok=True)
                    tmp_ud = str(TEMP_DIR / "edge-user-data")
                    print(f"[警告] 将回退为临时用户数据目录启动：{tmp_ud}")
                    opts3 = _make_edge_options(tmp_ud, "Default")
                    drv = webdriver.Edge(options=opts3)
                    drv.implicitly_wait(2)
                    return drv
                except Exception as e3:
                    print(f"[错误] 使用临时用户数据目录启动也失败：{e3}")
                    raise

        # 未识别的异常，直接抛出
        raise e

    if service is None:
        print("[警告] 未找到 driver/msedgedriver.exe，将尝试自动匹配驱动（Selenium Manager）")
        try:
            driver = webdriver.Edge(options=opts)
        except (SessionNotCreatedException, WebDriverException) as e:
            driver = _try_attach_mode_on_failure(e)
    else:
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
                # 交由统一回退处理（包含“用户数据目录被占用”）
                driver = _try_attach_mode_on_failure(e)
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
    print(f"[信息] 启动 Edge 远程调试实例：{exe}，端口 {port}，用户数据 {user_data_dir} / {profile_dir}")
    subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
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
    """等待页面出现“添加规格类型”按钮。"""
    X_ADD_SPEC_BUTTON = (
        By.XPATH,
        "//button[contains(@class,'ecom-g-btn-dashed')]//span[contains(normalize-space(.),'添加规格类型')]/ancestor::button",
    )
    deadline = time.time() + max_total_seconds
    has_prompted = False
    while time.time() < deadline:
        try:
            WebDriverWait(driver, timeout).until(EC.presence_of_element_located(X_ADD_SPEC_BUTTON))
            print("[信息] 页面已加载，检测到“添加规格类型”按钮")
            # 页面就绪后，提前滚动到“添加规格类型”按钮附近，便于后续操作区域可视
            try:
                btn = driver.find_element(*X_ADD_SPEC_BUTTON)
                driver.execute_script(
                    """
                    (function(el){
                        try {
                            var rect = el.getBoundingClientRect();
                            var top = rect.top + (window.pageYOffset || document.documentElement.scrollTop || document.body.scrollTop || 0) - 140;
                            if (top < 0) top = 0;
                            window.scrollTo(0, top);
                        } catch(e) { try { el.scrollIntoView({block:'center'}); } catch(e2) {} }
                    })(arguments[0]);
                    """,
                    btn,
                )
            except Exception:
                pass
            return
        except Exception:
            if not has_prompted:
                print("[提示] 检测到登录状态可能失效或页面尚未就绪，请在浏览器完成登录；程序会持续等待，直到检测到“添加规格类型”按钮或超时（最多10分钟）…")
                has_prompted = True
            time.sleep(2)
    raise RuntimeError("长时间未检测到“添加规格类型”按钮，可能未完成登录或页面异常")


def navigate_to_url(driver: webdriver.Edge, url: str, wait_seconds: int = 5) -> None:
    """稳健导航到指定 URL：driver.get / window.open / location.href 多方案回退。"""
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



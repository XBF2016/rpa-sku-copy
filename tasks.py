from robocorp.tasks import task
import subprocess
import time
import os
import shutil
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.edge.service import Service as EdgeService


def _read_browser_path() -> str:
    """从 browser.txt 读取 Edge 可执行文件路径；若为空则回退为 'msedge.exe'。"""
    path_file = Path(__file__).with_name("browser.txt")
    try:
        content = path_file.read_text(encoding="utf-8").strip().strip('"')
        if content:
            return content
    except Exception:
        pass
    return "msedge.exe"


def _is_process_running(image_name: str) -> bool:
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


def _kill_edge_processes() -> None:
    """强制结束 Edge 相关进程（msedge.exe、msedgewebview2.exe）。"""
    for image in ("msedge.exe", "msedgewebview2.exe"):
        try:
            subprocess.run(["taskkill", "/IM", image, "/F"], capture_output=True, text=True)
        except Exception:
            # 忽略错误以保证幂等性
            pass


def _kill_driver_processes() -> None:
    """强制结束 EdgeDriver 相关进程（msedgedriver.exe）。"""
    for image in ("msedgedriver.exe",):
        try:
            subprocess.run(["taskkill", "/IM", image, "/F"], capture_output=True, text=True)
        except Exception:
            pass


def _pid_running(pid: int) -> bool:
    """判断指定 PID 的进程是否存在。"""
    try:
        if pid <= 0:
            return False
        result = subprocess.run(
            ["tasklist", "/FI", f"PID eq {pid}"], capture_output=True, text=True, check=False
        )
        out = (result.stdout or "")
        # 当找不到时输出中会包含 "No tasks are running" 或不包含 PID
        return str(pid) in out
    except Exception:
        return False


def _acquire_single_instance_lock(name: str) -> Path:
    """创建单实例锁，避免多个任务并发运行。
    锁文件位于系统临时目录：<tmp>/rpa-sku-copy-<name>.lock
    若锁已存在且对应 PID 仍在运行，则抛出异常。
    若锁已存在但 PID 不存在，则清理后重新加锁。
    """
    import tempfile
    lock_path = Path(tempfile.gettempdir()) / f"rpa-sku-copy-{name}.lock"
    try:
        if lock_path.exists():
            try:
                pid_text = lock_path.read_text(encoding="utf-8").strip()
                pid = int(pid_text) if pid_text.isdigit() else -1
            except Exception:
                pid = -1
            if _pid_running(pid):
                raise RuntimeError(f"已有同类任务正在运行（PID={pid}），请稍后再试或手动结束该进程。")
            # 清理僵尸锁
            try:
                lock_path.unlink(missing_ok=True)
            except Exception:
                pass
        # 写入当前 PID
        lock_path.write_text(str(os.getpid()), encoding="utf-8")
        return lock_path
    except Exception as e:
        raise RuntimeError(f"创建单实例锁失败：{e}")


def _release_single_instance_lock(lock_path: Path) -> None:
    """释放单实例锁。"""
    try:
        if lock_path and lock_path.exists():
            lock_path.unlink(missing_ok=True)
    except Exception:
        pass


def _find_msedgedriver_path() -> str | None:
    """在本机查找 msedgedriver.exe 的路径，尽量避免联网下载。
    查找顺序：
      1) 环境变量 MSEDGEDRIVER 指定的路径
      2) 项目根目录下的 msedgedriver.exe（与 tasks.py 同级）
      3) 与 browser.txt 中的 msedge.exe 同目录下的 msedgedriver.exe
      4) 常见安装目录：
         - C:\\Program Files\\Microsoft\\Edge\\Application\\msedgedriver.exe
         - C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedgedriver.exe
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

    # 2) 项目根目录
    proj_driver = Path(__file__).with_name("msedgedriver.exe")
    if proj_driver.exists():
        return str(proj_driver)

    # 3) 与 msedge.exe 同目录
    edge_exe = _read_browser_path()
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


def _read_product_url() -> str:
    """从 product-url.txt 读取商品链接（取第一条非空行）。"""
    path_file = Path(__file__).with_name("product-url.txt")
    if not path_file.exists():
        raise ValueError("未找到 product-url.txt，请在项目根目录提供该文件")
    # 读取第一条非空行
    for line in path_file.read_text(encoding="utf-8").splitlines():
        url = line.strip()
        if url:
            if url.startswith("http://") or url.startswith("https://"):
                return url
            raise ValueError("product-url.txt 中的第一条非空内容不是有效的 http/https 链接")
    raise ValueError("product-url.txt 中没有可用的链接内容")


@task
def open_edge_logged_in():
    """
    关闭任何正在运行的 Microsoft Edge 实例，然后使用 Default 配置文件启动 Edge。
    Default 配置文件通常会保留用户的登录状态。
    浏览器可执行文件路径从 'browser.txt' 读取。
    """
    print("[步骤] 检查 Edge 是否正在运行...")
    if _is_process_running("msedge.exe"):
        print("[信息] 检测到 Edge 正在运行，尝试关闭所有 Edge 相关进程...")
        _kill_edge_processes()
        time.sleep(1.5)
    else:
        print("[信息] 未检测到 Edge 实例。")

    exe_path = _read_browser_path()
    exe = exe_path
    if not os.path.exists(exe_path) and not shutil.which(exe_path):
        print(f"[警告] 未在 '{exe_path}' 找到 Edge 可执行文件，回退为 'msedge.exe'。")
        exe = "msedge.exe"

    # 使用 Default 配置文件以复用登录状态；打开空白页并最大化窗口
    args = [exe, "--profile-directory=Default", "--start-maximized", "about:blank"]
    print(f"[步骤] 启动 Edge: {' '.join(args)}")
    try:
        subprocess.Popen(args, shell=False)
        print("[完成] 已使用 Default 配置文件启动 Edge。若你的 Default 配置已登录，\n"
              "       本次会话将保持登录状态。")
    except Exception as e:
        print(f"[错误] 启动 Edge 失败: {e}")
        raise


@task
def open_product_page():
    """
    在不强制关闭 Edge 的情况下，读取 product-url.txt，并在 Edge 中打开该商品页面。
    若 Edge 未在运行，则以 Default 配置文件启动并直接打开目标链接。
    """
    url = _read_product_url()
    print(f"[步骤] 读取到商品链接: {url}")

    exe_path = _read_browser_path()
    exe = exe_path if os.path.exists(exe_path) or shutil.which(exe_path) else "msedge.exe"
    if exe != exe_path:
        print(f"[警告] 未在 '{exe_path}' 找到 Edge 可执行文件，回退为 'msedge.exe'。")

    if _is_process_running("msedge.exe"):
        # Edge 已在运行：再次调用可执行程序并传入 URL，会在现有实例中新开标签页
        args = [exe, "--profile-directory=Default", url]
        print(f"[步骤] Edge 已运行，尝试在现有实例中新开标签并访问商品页面...")
    else:
        # Edge 未运行：启动并直接打开该 URL
        args = [exe, "--profile-directory=Default", "--start-maximized", url]
        print(f"[步骤] Edge 未运行，尝试启动并直接访问商品页面...")

    print(f"[信息] 调用命令: {' '.join(args)}")
    try:
        subprocess.Popen(args, shell=False)
        print("[完成] 已在 Edge 中打开商品页面。")
    except Exception as e:
        print(f"[错误] 打开商品页面失败: {e}")
        raise


@task
def extract_sku_dimensions():
    """
    打开商品页，自动识别所有 SKU 维度及其下的所有选项，并写入 log/sku维度及选项.log。
    首先关闭现有Edge进程，然后使用用户登录态重新启动Edge。
    定位规则严格依据你提供的元素结构与类名：
      - 维度容器：.skuItem--Z2AJB9Ew
      - 维度名称：.ItemLabel--psS1SOyC > span.f-els-2
      - 选项容器：.skuValueWrap--aEfxuhNr .content--DIGuLqdf .valueItem--smR4pNt4
      - 选项文本：选项容器内的 span.f-els-1 的 title 或文本
    """
    url = _read_product_url()
    print(f"[步骤] 读取到商品链接: {url}")

    # 首先关闭现有的Edge与EdgeDriver进程以避免冲突
    print("[步骤] 检查并关闭现有的 Edge 进程...")
    if _is_process_running("msedge.exe"):
        print("[信息] 检测到 Edge 正在运行，关闭所有 Edge 相关进程...")
        _kill_edge_processes()
        time.sleep(2)  # 等待进程完全关闭
    else:
        print("[信息] 未检测到 Edge 实例。")
    print("[步骤] 关闭可能残留的 EdgeDriver 进程...")
    _kill_driver_processes()

    # 获取用户数据目录以保持登录态
    import os
    user_data_dir = os.path.expanduser("~\\AppData\\Local\\Microsoft\\Edge\\User Data")
    
    print("[步骤] 初始化 Edge WebDriver（使用用户登录态）...")
    options = EdgeOptions()
    options.add_argument("--start-maximized")
    options.add_argument(f"--user-data-dir={user_data_dir}")  # 使用用户数据目录保持登录态
    options.add_argument("--profile-directory=Default")  # 使用默认配置文件
    # 避免一些常见的安全限制
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)

    # 使用本地 msedgedriver.exe，避免联网下载
    driver_path = _find_msedgedriver_path()
    if not driver_path:
        print("[错误] 未找到 msedgedriver.exe。请下载与本机 Edge 版本匹配的 EdgeDriver，\n"
              "       并将 msedgedriver.exe 放到项目根目录，或设置环境变量 MSEDGEDRIVER 指向其目录/文件。")
        raise RuntimeError("未找到本地 EdgeDriver")

    print(f"[信息] 使用本地 EdgeDriver: {driver_path}")
    print(f"[信息] 使用用户数据目录: {user_data_dir}")
    try:
        driver = webdriver.Edge(service=EdgeService(executable_path=driver_path), options=options)
    except Exception as e:
        print(f"[错误] 初始化 Edge WebDriver 失败: {e}")
        raise

    # 单实例锁，避免并发写日志
    lock_path = _acquire_single_instance_lock("extract")
    try:
        print("[步骤] 打开商品页面...")
        driver.get(url)

        wait = WebDriverWait(driver, 25)
        # 等待至少出现一个维度块
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".skuItem--Z2AJB9Ew")))

        # 收集所有维度
        sku_items = driver.find_elements(By.CSS_SELECTOR, ".skuItem--Z2AJB9Ew")
        print(f"[信息] 检测到维度数量: {len(sku_items)}")

        results: list[str] = []

        for idx, item in enumerate(sku_items, start=1):
            # 维度名
            dim_name = ""
            try:
                label_span = item.find_element(By.CSS_SELECTOR, ".ItemLabel--psS1SOyC span.f-els-2")
                dim_name = label_span.get_attribute("title") or label_span.text or ""
                dim_name = dim_name.strip()
            except Exception:
                dim_name = f"未命名维度#{idx}"

            if not dim_name:
                dim_name = f"未命名维度#{idx}"

            # 选项
            option_nodes = item.find_elements(By.CSS_SELECTOR, ".skuValueWrap--aEfxuhNr .content--DIGuLqdf .valueItem--smR4pNt4")
            option_texts: list[str] = []
            for opt in option_nodes:
                text = ""
                try:
                    span = opt.find_element(By.CSS_SELECTOR, "span.f-els-1")
                    text = span.get_attribute("title") or span.text or ""
                except Exception:
                    # 兜底：取整个节点的可见文本
                    text = opt.text or ""
                text = (text or "").strip()
                if text:
                    option_texts.append(text)

            # 写入内存缓存
            results.append(f"维度：{dim_name}")
            if option_texts:
                for t in option_texts:
                    results.append(f"  - {t}")
            else:
                results.append("  - （无选项或未能识别）")

        # 输出到日志文件
        log_dir = Path(__file__).with_name("log")
        log_dir.mkdir(parents=True, exist_ok=True)
        log_path = log_dir / "sku维度及选项.log"
        print(f"[步骤] 将识别结果写入: {log_path}")
        content = "\n".join(results) + "\n"
        log_path.write_text(content, encoding="utf-8")
        print("[完成] SKU 维度与选项已写入日志文件。")

    finally:
        try:
            driver.quit()
        except Exception:
            pass
        # 释放锁并兜底清理驱动进程
        _release_single_instance_lock(lock_path)
        _kill_driver_processes()


@task
def traverse_all_sku_combinations():
    """
    自动遍历所有SKU维度组合，获取每个组合的价格信息，并写入markdown表格。
    使用防检测机制：随机延迟、模拟人类行为、控制点击速度。
    总共需要遍历 16×4×6=384个组合。
    """
    import random
    import itertools
    from selenium.webdriver.common.action_chains import ActionChains
    
    url = _read_product_url()
    print(f"[步骤] 读取到商品链接: {url}")

    # 关闭现有Edge与EdgeDriver进程
    print("[步骤] 检查并关闭现有的 Edge 进程...")
    if _is_process_running("msedge.exe"):
        print("[信息] 检测到 Edge 正在运行，关闭所有 Edge 相关进程...")
        _kill_edge_processes()
        time.sleep(2)
    print("[步骤] 关闭可能残留的 EdgeDriver 进程...")
    _kill_driver_processes()

    # 获取用户数据目录以保持登录态
    user_data_dir = os.path.expanduser("~\\AppData\\Local\\Microsoft\\Edge\\User Data")
    
    print("[步骤] 初始化 Edge WebDriver（使用用户登录态）...")
    options = EdgeOptions()
    options.add_argument("--start-maximized")
    options.add_argument(f"--user-data-dir={user_data_dir}")
    options.add_argument("--profile-directory=Default")
    # 防检测设置
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    
    driver_path = _find_msedgedriver_path()
    if not driver_path:
        print("[错误] 未找到 msedgedriver.exe")
        raise RuntimeError("未找到本地 EdgeDriver")

    print(f"[信息] 使用本地 EdgeDriver: {driver_path}")
    driver = webdriver.Edge(service=EdgeService(executable_path=driver_path), options=options)
    
    # 设置较短的隐式等待和脚本执行器
    driver.implicitly_wait(1)  # 从10秒减少到1秒
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    
    try:
        print("[步骤] 打开商品页面...")
        driver.get(url)
        time.sleep(random.uniform(3, 5))  # 随机等待页面加载
        
        wait = WebDriverWait(driver, 30)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".skuItem--Z2AJB9Ew")))
        
        # 解析所有维度和选项
        print("[步骤] 解析SKU维度和选项...")
        sku_dimensions = []
        sku_items = driver.find_elements(By.CSS_SELECTOR, ".skuItem--Z2AJB9Ew")
        
        for idx, item in enumerate(sku_items, start=1):
            # 获取维度名称
            try:
                dim_label = item.find_element(By.CSS_SELECTOR, ".ItemLabel--psS1SOyC span.f-els-2")
                dim_name = dim_label.get_attribute("title") or dim_label.text or f"维度{idx}"
            except Exception:
                dim_name = f"维度{idx}"
            
            # 获取该维度下的所有选项
            options = []
            option_elements = item.find_elements(By.CSS_SELECTOR, ".valueItem--smR4pNt4")
            
            for opt_elem in option_elements:
                try:
                    # 检查选项是否可用
                    if opt_elem.get_attribute("data-disabled") == "true":
                        continue
                        
                    data_vid = opt_elem.get_attribute("data-vid")
                    span = opt_elem.find_element(By.CSS_SELECTOR, "span.f-els-1")
                    option_text = span.get_attribute("title") or span.text or ""
                    
                    if data_vid and option_text:
                        options.append({
                            "vid": data_vid,
                            "text": option_text.strip(),
                            "element": opt_elem
                        })
                except Exception:
                    continue
            
            if options:
                sku_dimensions.append({
                    "name": dim_name.strip(),
                    "options": options
                })
        
        total_combinations = 1
        for dim in sku_dimensions:
            total_combinations *= len(dim["options"])
        
        print(f"[信息] 检测到 {len(sku_dimensions)} 个维度，总共 {total_combinations} 个组合")
        for i, dim in enumerate(sku_dimensions):
            print(f"  维度{i+1}: {dim['name']} ({len(dim['options'])}个选项)")
        
        # 生成所有可能的组合
        print("[步骤] 开始遍历所有SKU组合...")
        combinations = list(itertools.product(*[dim["options"] for dim in sku_dimensions]))
        
        results = []
        success_count = 0
        
        for combo_idx, combination in enumerate(combinations, 1):
            try:
                print(f"\n[进度] 处理组合 {combo_idx}/{len(combinations)}")
                
                # 构建组合描述
                combo_text = []
                for i, option in enumerate(combination):
                    combo_text.append(f"{sku_dimensions[i]['name']}: {option['text']}")
                print(f"[组合] {' | '.join(combo_text)}")
                
                # 记录组合信息到log文件
                combo_info = f"{combination[0]['text']}---{combination[1]['text']}---{combination[2]['text']}"
                _append_to_log(f"正在处理组合 {combo_idx}/{len(combinations)}: {combo_info}")
                
                # 依次点击每个维度的选项
                start_time = time.time()
                for dim_idx, option in enumerate(combination):
                    try:
                        # 重新找到元素（防止DOM更新导致的stale element问题）
                        sku_items = driver.find_elements(By.CSS_SELECTOR, ".skuItem--Z2AJB9Ew")
                        current_dim = sku_items[dim_idx]
                        
                        # 找到对应的选项元素
                        option_selector = f'.valueItem--smR4pNt4[data-vid="{option["vid"]}"]'
                        option_element = current_dim.find_element(By.CSS_SELECTOR, option_selector)
                        
                        # 使用JavaScript点击避免拦截问题
                        driver.execute_script("arguments[0].click();", option_element)
                        
                        # 很短的延迟
                        time.sleep(0.1)
                        
                    except Exception as e:
                        print(f"[警告] 点击维度{dim_idx+1}选项失败: {e}")
                        continue
                
                # 等待价格更新（较短时间）
                time.sleep(0.3)
                
                # 快速获取价格信息（优化版）
                price = "未获取到价格"
                try:
                    # 使用更高效的方法：先尝试最常见的选择器
                    main_selectors = [
                        ".price--yeTcvSlD .number--ZQ6CbUNc",  # 淘宝最常见价格选择器
                        ".tm-price-current",
                        "[class*='price'] [class*='number']"
                    ]
                    
                    # 临时设置更短的隐式等待
                    driver.implicitly_wait(0.5)
                    
                    for selector in main_selectors:
                        try:
                            price_element = driver.find_element(By.CSS_SELECTOR, selector)
                            price_text = price_element.text.strip()
                            if price_text and price_text != "":
                                price = price_text
                                break
                        except Exception:
                            continue
                    
                    # 如果还没找到，快速尝试通用方法
                    if price == "未获取到价格":
                        try:
                            price_elements = driver.find_elements(By.XPATH, "//*[contains(text(), '¥') or contains(text(), '￥')]")
                            for elem in price_elements[:2]:  # 只检查前2个
                                text = elem.text.strip()
                                if text and ('¥' in text or '￥' in text) and len(text) < 20:  # 限制长度避免获取到描述文本
                                    price = text
                                    break
                        except Exception:
                            pass
                    
                    # 恢复隐式等待
                    driver.implicitly_wait(1)
                
                except Exception as e:
                    print(f"[警告] 获取价格失败: {e}")
                    # 恢复隐式等待
                    driver.implicitly_wait(1)
                
                # 保存结果
                result_row = []
                for option in combination:
                    result_row.append(option['text'])
                result_row.append(price)
                results.append(result_row)
                
                success_count += 1
                print(f"[成功] 价格: {price}")
                
                # 记录成功结果到log文件
                full_combo_info = f"{combo_info}---{price}"
                _append_to_log(f"完成组合 {combo_idx}: {full_combo_info}")
                
                # 控制每个组合处理时间约1秒（减少不必要延迟）
                elapsed_time = time.time() - start_time
                if elapsed_time < 0.8:  # 从1秒减少到0.8秒，更快的处理速度
                    time.sleep(0.8 - elapsed_time)
                
                # 每处理一定数量的组合就保存一次结果
                if combo_idx % 50 == 0:
                    _append_to_log(f"中间保存: 已处理 {combo_idx} 个组合")
                    print(f"[信息] 已处理 {combo_idx} 个组合，中间保存完成")
                
                # 防检测：每50个组合休息一下（减少频率）
                if combo_idx % 50 == 0:
                    delay = random.uniform(1, 2)
                    _append_to_log(f"防检测休息 {delay:.1f} 秒...")
                    print(f"[防检测] 休息 {delay:.1f} 秒...")
                    time.sleep(delay)
                
            except Exception as e:
                print(f"[错误] 处理组合 {combo_idx} 失败: {e}")
                continue
        
        # 保存最终结果
        print(f"\n[步骤] 遍历完成！成功处理 {success_count}/{len(combinations)} 个组合")
        print("[完成] 所有SKU组合遍历完成，结果已保存到 sku维度及选项.md")
        
    finally:
        try:
            driver.quit()
        except Exception:
            pass




def _append_to_log(message):
    """向log文件追加信息"""
    import datetime
    log_dir = Path(__file__).with_name("log")
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / "sku维度及选项.log"
    
    # 添加时间戳
    timestamp = datetime.datetime.now().strftime("%H:%M:%S")
    log_line = f"[{timestamp}] {message}\n"
    
    # 追加写入，处理编码问题
    try:
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(log_line)
            f.flush()  # 确保立即写入
    except UnicodeEncodeError:
        # 如果出现编码问题，清理特殊字符
        clean_message = message.encode('utf-8', errors='ignore').decode('utf-8')
        clean_log_line = f"[{timestamp}] {clean_message}\n"
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(clean_log_line)
            f.flush()

from robocorp.tasks import task
import subprocess
import time
import os
import shutil
from pathlib import Path
from dataclasses import dataclass
from typing import List, Tuple
import random
import itertools
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.edge.service import Service as EdgeService


# 统一的页面元素选择器常量，便于维护与复用
SKU_ITEM_SELECTOR = ".skuItem--Z2AJB9Ew"
SKU_OPTION_SELECTOR = ".valueItem--smR4pNt4"
DIM_LABEL_SELECTOR = ".ItemLabel--psS1SOyC span.f-els-2"
OPTION_TEXT_SELECTOR = "span.f-els-1"

# 价格相关选择器
PRICE_MAIN_TEXT = ".highlightPrice--asfw5V1e .text--jyiUrkMu"
PRICE_SYMBOL = ".highlightPrice--asfw5V1e .symbol--ZqZXkLDL"
ORIG_PRICE_TEXTS = ".subPrice--empS5uv8 .text--jyiUrkMu"
PRICE_ALT_SELECTORS = [
    ".beltPrice--i5j_t2w4 .text--jyiUrkMu",
    ".price--yeTcvSlD .number--ZQ6CbUNc",
    ".tm-price-current",
    "[class*='price'] [class*='number']",
    "[class*='highlightPrice'] [class*='text']",
]


# 控制台/日志文本规范化（解决 Windows 控制台 GBK 下 '¥' 无法编码的问题）
def _normalize_price_text(txt: str) -> str:
    try:
        if not isinstance(txt, str):
            txt = str(txt)
    except Exception:
        txt = ""
    # 将 '¥'(U+00A5) 统一替换为 '￥'(U+FFE5)
    return (txt or "").replace("¥", "￥").strip()


# 数据模型：更通用、更可读
@dataclass(frozen=True)
class SkuOption:
    vid: str
    text: str


@dataclass
class SkuDimension:
    name: str
    options: List[SkuOption]

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
    """从 product-url.txt 读取商品链接（取第一条有效行）。
    规则：
      - 忽略空行
      - 忽略以 '#' 开头的注释行
      - 取第一条以 http/https 开头的链接
    """
    path_file = Path(__file__).with_name("product-url.txt")
    if not path_file.exists():
        raise ValueError("未找到 product-url.txt，请在项目根目录提供该文件")
    # 顺序读取，跳过空行与以 # 开头的注释行，返回第一条有效链接
    for line in path_file.read_text(encoding="utf-8").splitlines():
        url = line.strip()
        # 跳过空行
        if not url:
            continue
        # 跳过注释行（以 # 开头）
        if url.startswith('#'):
            continue
        # 校验链接协议头
        if url.startswith("http://") or url.startswith("https://"):
            return url
        # 如果遇到的第一条非空且非注释内容不是链接，则直接报错
        raise ValueError("product-url.txt 中存在非注释的第一条内容不是有效的 http/https 链接")
    raise ValueError("product-url.txt 中没有可用的链接内容（已忽略空行与注释行）")






def _prepare_clean_edge_state() -> None:
    """准备干净的 Edge 运行环境，确保不受残留进程影响。"""
    print("[步骤] 检查并关闭现有的 Edge 进程...")
    if _is_process_running("msedge.exe"):
        print("[信息] 检测到 Edge 正在运行，关闭所有 Edge 相关进程...")
        _kill_edge_processes()
        time.sleep(2)
    print("[步骤] 关闭可能残留的 EdgeDriver 进程...")
    _kill_driver_processes()


def _init_edge_driver() -> webdriver.Edge:
    """初始化 Edge WebDriver（使用本地 driver 与用户登录态）。"""
    user_data_dir = os.path.expanduser("~\\AppData\\Local\\Microsoft\\Edge\\User Data")

    print("[步骤] 初始化 Edge WebDriver（使用用户登录态）...")
    options = EdgeOptions()
    options.add_argument("--start-maximized")
    options.add_argument(f"--user-data-dir={user_data_dir}")
    options.add_argument("--profile-directory=Default")
    # 防检测设置
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])  # 排除日志开关
    options.add_experimental_option('useAutomationExtension', False)
    # 降低浏览器日志级别，仅保留致命错误；并关闭日志输出、禁用QUIC，忽略证书错误
    options.add_argument("--log-level=3")
    options.add_argument("--disable-logging")
    options.add_argument("--disable-quic")
    options.add_argument("--ignore-certificate-errors")
    options.set_capability("acceptInsecureCerts", True)

    driver_path = _find_msedgedriver_path()
    if not driver_path:
        print("[错误] 未找到 msedgedriver.exe")
        raise RuntimeError("未找到本地 EdgeDriver")

    print(f"[信息] 使用本地 EdgeDriver: {driver_path}")
    try:
        service = EdgeService(executable_path=driver_path, log_output=subprocess.DEVNULL)
    except TypeError:
        service = EdgeService(executable_path=driver_path)

    driver = webdriver.Edge(service=service, options=options)
    driver.implicitly_wait(1)
    # 移除 webdriver 痕迹
    try:
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    except Exception:
        pass
    return driver


def _open_product_page(driver: webdriver.Edge, url: str, wait_timeout: int = 30) -> None:
    """打开商品页并等待 SKU 区域出现。"""
    print("[步骤] 打开商品页面...")
    driver.get(url)
    # 页面打开后做一个短随机等待（防检测/渲染稳定）
    time.sleep(random.uniform(0.6, 1.2))
    WebDriverWait(driver, wait_timeout).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, SKU_ITEM_SELECTOR))
    )


def _parse_sku_dimensions(driver: webdriver.Edge) -> List[SkuDimension]:
    """解析页面上的所有 SKU 维度与选项。"""
    print("[步骤] 解析SKU维度和选项...")
    sku_dimensions: List[SkuDimension] = []
    sku_items = driver.find_elements(By.CSS_SELECTOR, SKU_ITEM_SELECTOR)

    for idx, item in enumerate(sku_items, start=1):
        # 维度名
        try:
            dim_label = item.find_element(By.CSS_SELECTOR, DIM_LABEL_SELECTOR)
            dim_name = dim_label.get_attribute("title") or dim_label.text or f"维度{idx}"
        except Exception:
            dim_name = f"维度{idx}"

        # 选项集合
        options: List[SkuOption] = []
        option_elements = item.find_elements(By.CSS_SELECTOR, SKU_OPTION_SELECTOR)
        for opt_elem in option_elements:
            try:
                # 跳过不可用
                if (opt_elem.get_attribute("data-disabled") or "").lower() == "true":
                    continue
                data_vid = opt_elem.get_attribute("data-vid")
                span = opt_elem.find_element(By.CSS_SELECTOR, OPTION_TEXT_SELECTOR)
                option_text = (span.get_attribute("title") or span.text or "").strip()
                if data_vid and option_text:
                    options.append(SkuOption(vid=data_vid, text=option_text))
            except Exception:
                continue

        if options:
            sku_dimensions.append(SkuDimension(name=dim_name.strip(), options=options))

    return sku_dimensions


def _is_selected_element(elem) -> bool:
    """判断一个选项元素是否处于选中状态。"""
    try:
        cls = (elem.get_attribute("class") or "")
        if ("selected" in cls) or ("Selected" in cls) or ("active" in cls) or ("checked" in cls):
            return True
        if (elem.get_attribute("aria-checked") or "").lower() == "true":
            return True
        if (elem.get_attribute("data-selected") or "").lower() == "true":
            return True
    except Exception:
        pass
    return False


def _ensure_combination_selected(
    driver: webdriver.Edge,
    combination: List[SkuOption],
    last_selected_vids: List[str] | None = None,
) -> List[str]:
    """按需点击组合中的 SKU 选项：仅对发生变化的维度执行点击；随后做一次快速校验，必要时补点。
    返回本次目标组合的 vid 列表，供下次迭代复用，减少无效点击。
    """
    # 计算需要变更的维度索引集合
    need_change_indices = list(range(len(combination)))
    if last_selected_vids and len(last_selected_vids) == len(combination):
        need_change_indices = [
            i for i, opt in enumerate(combination) if last_selected_vids[i] != opt.vid
        ]

    # 如果没有变化，直接返回
    if not need_change_indices:
        return [opt.vid for opt in combination]

    # 首轮：只点击有变化的维度
    for dim_idx in need_change_indices:
        option = combination[dim_idx]
        try:
            sku_items = driver.find_elements(By.CSS_SELECTOR, SKU_ITEM_SELECTOR)
            current_dim = sku_items[dim_idx]
            option_selector = f'{SKU_OPTION_SELECTOR}[data-vid="{option.vid}"]'
            option_element = current_dim.find_element(By.CSS_SELECTOR, option_selector)
            if not _is_selected_element(option_element):
                driver.execute_script("arguments[0].click();", option_element)
                # 极短抖动，给页面联动微时间
                time.sleep(0.03)
        except Exception as e:
            print(f"[警告] 点击维度{dim_idx+1}选项失败: {e}")
            continue

    # 二次：对全部维度做一次快速校验与补点，避免上层维度变化导致下层被反选
    try:
        sku_items = driver.find_elements(By.CSS_SELECTOR, SKU_ITEM_SELECTOR)
        for dim_idx, option in enumerate(combination):
            current_dim = sku_items[dim_idx]
            option_selector = f'{SKU_OPTION_SELECTOR}[data-vid="{option.vid}"]'
            option_element = current_dim.find_element(By.CSS_SELECTOR, option_selector)
            if not _is_selected_element(option_element):
                driver.execute_script("arguments[0].click();", option_element)
                time.sleep(0.03)
    except Exception:
        pass

    return [opt.vid for opt in combination]


def _get_price_text(driver: webdriver.Edge) -> str:
    """获取当前所选组合的价格文本（JS一次性查询 + 短轮询，极限压缩等待时间）。"""
    # 通过 execute_script 的参数传入备用选择器，避免引号转义问题
    js = (
        "return (function(alts){\n"
        f"  var mainText = document.querySelector('{PRICE_MAIN_TEXT}');\n"
        f"  var symbolEl = document.querySelector('{PRICE_SYMBOL}');\n"
        "  if (mainText && mainText.textContent) {\n"
        "    var sym = symbolEl && symbolEl.textContent ? symbolEl.textContent.trim() : '¥';\n"
        "    var txt = mainText.textContent.trim();\n"
        "    if (txt) return sym + txt;\n"
        "  }\n"
        "  alts = Array.isArray(alts) ? alts : [];\n"
        "  for (var i=0; i<alts.length; i++){\n"
        "    var el = document.querySelector(alts[i]);\n"
        "    if (el && el.textContent){\n"
        "      var t = el.textContent.trim();\n"
        "      if (t){\n"
        "        if (t.indexOf('¥') !== -1 || t.indexOf('￥') !== -1) return t;\n"
        "        return '¥' + t;\n"
        "      }\n"
        "    }\n"
        "  }\n"
        "  var belt = document.querySelector('.beltPrice--i5j_t2w4');\n"
        "  if (belt){\n"
        "    var nodes = belt.querySelectorAll(\".text--jyiUrkMu, .number--ZQ6CbUNc, [class*='number'], [class*='text']\");\n"
        "    for (var j=0; j<nodes.length; j++){\n"
        "      var tt = (nodes[j].textContent || '').trim();\n"
        "      if (tt && /\\d/.test(tt)){\n"
        "        if (tt.indexOf('¥') !== -1 || tt.indexOf('￥') !== -1) return tt;\n"
        "        return '¥' + tt;\n"
        "      }\n"
        "    }\n"
        "  }\n"
        "  return '';\n"
        "})(arguments[0]);"
    )

    # 最多短轮询 ~300ms，提高对价格异步刷新的兼容
    end_time = time.perf_counter() + 0.3
    last = ""
    while time.perf_counter() < end_time:
        try:
            price = (driver.execute_script(js, PRICE_ALT_SELECTORS) or "").strip()
            if price and any(ch.isdigit() for ch in price):
                return _normalize_price_text(price)
            last = price
        except Exception:
            pass
        time.sleep(0.06)
    return _normalize_price_text(last) or "未获取到价格"



@task
def traverse_all_sku_combinations():
    """
    自动遍历所有SKU维度组合，获取每个组合的价格信息，并写入markdown表格。
    使用防检测机制：随机延迟、模拟人类行为、控制点击速度。
    """
    url = _read_product_url()
    print(f"[步骤] 读取到商品链接: {url}")

    # 准备浏览器环境并初始化驱动
    _prepare_clean_edge_state()
    driver = _init_edge_driver()

    try:
        # 打开页面并等待初始加载
        _open_product_page(driver, url, wait_timeout=30)

        # 解析维度
        sku_dimensions = _parse_sku_dimensions(driver)
        total_combinations = 1
        for dim in sku_dimensions:
            total_combinations *= len(dim.options)

        print(f"[信息] 检测到 {len(sku_dimensions)} 个维度，总共 {total_combinations} 个组合")
        for i, dim in enumerate(sku_dimensions):
            print(f"  维度{i+1}: {dim.name} ({len(dim.options)}个选项)")

        print("[步骤] 开始遍历所有SKU组合...")
        combinations: List[Tuple[SkuOption, ...]] = list(itertools.product(*[d.options for d in sku_dimensions]))
        # 支持通过环境变量限制前 N 个组合用于快速自测
        try:
            max_combos_str = os.environ.get("MAX_COMBOS", "").strip()
            if max_combos_str:
                max_combos = int(max_combos_str)
                if max_combos > 0 and max_combos < len(combinations):
                    print(f"[信息] 仅执行前 {max_combos} 个组合用于快速验证（通过 MAX_COMBOS 控制）")
                    combinations = combinations[:max_combos]
        except Exception:
            pass

        results: List[List[str]] = []
        success_count = 0
        last_selected_vids: List[str] = []

        for combo_idx, combination in enumerate(combinations, 1):
            try:
                print(f"\n[进度] 处理组合 {combo_idx}/{len(combinations)}")
                combo_text_parts = [f"{sku_dimensions[i].name}: {opt.text}" for i, opt in enumerate(combination)]
                print(f"[组合] {' | '.join(combo_text_parts)}")

                # 组合信息（通用拼接，不假定维度数量）
                combo_info = "---".join([opt.text for opt in combination])

                # 点击并校验
                start_time = time.time()
                last_selected_vids = _ensure_combination_selected(
                    driver, list(combination), last_selected_vids or None
                )
                price = _get_price_text(driver)
                # 再保险：统一控制台友好的货币符号
                price = _normalize_price_text(price)

                # 保存结果
                result_row = [opt.text for opt in combination] + [price]
                results.append(result_row)
                success_count += 1
                print(f"[成功] 价格: {price}")
                elapsed_time = time.time() - start_time
                print(f"[耗时] 本组合用时 {elapsed_time:.3f} 秒")

                full_combo_info = f"{combo_info}---{price}"
                _append_to_log(f"完成组合 {combo_idx}: {full_combo_info}")

                # 极短抖动，降低行为过于机械的特征，同时尽量不影响速度
                time.sleep(random.uniform(0.02, 0.06))

            except Exception as e:
                print(f"[错误] 处理组合 {combo_idx} 失败: {e}")
                continue

        print(f"\n[步骤] 遍历完成！成功处理 {success_count}/{len(combinations)} 个组合")

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


def main():
    """
    RPA任务统一入口函数
    按顺序执行所有RPA任务
    """
    try:
        print("=" * 50)
        print("开始执行RPA任务...")
        print("=" * 50)
        
        # 遍历所有SKU组合
        print("\n[额外步骤] 开始遍历所有SKU组合...")
        traverse_all_sku_combinations()
        print("✓ SKU组合遍历完成")
        
        print("\n" + "=" * 50)
        print("所有RPA任务执行完成！")
        print("=" * 50)
        
    except Exception as e:
        print(f"\n[错误] 执行过程中出现异常: {str(e)}")
        import traceback
        traceback.print_exc()
        return 1
    finally:
        # 确保所有进程都被正确清理
        _kill_driver_processes()
    
    return 0


if __name__ == "__main__":
    import sys
    sys.exit(main())

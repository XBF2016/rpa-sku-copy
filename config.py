# -*- coding: utf-8 -*-
"""
配置与路径常量、基础数据读取函数

- 路径常量：项目根目录、配置目录、驱动路径等
- 读取函数：浏览器路径、商品链接、规格配置（最简 YAML 解析）

说明：严格保持中文日志与报错，遵循项目原有规范；尽量少依赖第三方库。
"""
from __future__ import annotations

from pathlib import Path
from typing import Optional, List, Tuple


# -------------------------
# 路径与常量
# -------------------------
PROJECT_ROOT = Path(__file__).resolve().parent
CONF_DIR = PROJECT_ROOT / "conf"
DRIVER_PATH = PROJECT_ROOT / "driver" / "msedgedriver.exe"
BROWSER_PATH_FILE = CONF_DIR / "browser.txt"
TEMP_DIR = PROJECT_ROOT / "temp"
OUTPUT_DIR = PROJECT_ROOT / "output"

PRODUCT_URL_FILE_PRI = CONF_DIR / "douyin" / "product-url.txt"
PRODUCT_URL_FILE_FALLBACK = CONF_DIR / "product-url.txt"
SPECS_YAML_FILE = CONF_DIR / "规格.yml"


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


def read_spec_dimensions_with_options(yaml_path: Path) -> dict:
    """最简 YAML 解析：
    - 解析顶层键为维度名；其下缩进行若以 "- " 开头则作为该维度的选项。
    - 返回 { 维度: [选项1, 选项2, ...] } 的有序字典（按出现顺序）。
    """
    if not yaml_path.exists():
        raise RuntimeError(f"未找到规格配置文件: {yaml_path}")
    from collections import OrderedDict
    result = OrderedDict()
    cur_key = None
    try:
        with yaml_path.open("r", encoding="utf-8") as f:
            for raw in f:
                line = raw.rstrip("\n\r")
                if not line.strip():
                    continue
                if line.lstrip().startswith("#"):
                    continue
                # 顶层键：不以空格开头且以冒号结尾
                if not line.startswith(" ") and line.strip().endswith(":"):
                    cur_key = line.strip()[:-1].strip()
                    if cur_key:
                        result[cur_key] = []
                    continue
                # 子项：以连字符列表
                if cur_key and line.lstrip().startswith("-"):
                    # 允许任意缩进，但取"- "后的内容
                    dash_idx = line.find("-")
                    val = line[dash_idx + 1 :].strip()
                    if val:
                        result[cur_key].append(val)
    except Exception as e:
        raise RuntimeError(
            f"解析规格配置失败，请检查 conf/规格.yml 是否为 顶层键: 列表 的结构。错误: {e}"
        )
    if not result:
        raise RuntimeError("未从 conf/规格.yml 解析到任何维度，请确认内容格式")
    print(f"[信息] 将要创建的规格维度: {list(result.keys())}")
    return result


# -------------------------
# 新版：从 output/ 下最新子目录的 result.yml 读取规格（specs）
# -------------------------
def _find_latest_result_yml() -> Optional[Path]:
    """在 output/ 的一级子目录中查找最新的 result.yml。

    按文件修改时间选择最近的那个：output/<任意子目录>/result.yml
    若未找到，返回 None。
    """
    try:
        if not OUTPUT_DIR.exists():
            return None
        latest_file: Optional[Path] = None
        latest_mtime: float = -1.0
        for child in OUTPUT_DIR.iterdir():
            if not child.is_dir():
                continue
            cand = child / "result.yml"
            if cand.exists():
                try:
                    mtime = cand.stat().st_mtime
                except Exception:
                    mtime = 0.0
                if mtime > latest_mtime:
                    latest_mtime = mtime
                    latest_file = cand
        return latest_file
    except Exception:
        return None


def _parse_specs_from_result_yml(result_path: Path) -> dict:
    """最简 YAML 解析（仅解析 specs 段）：
    - 仅识别如下层级结构：
      specs:           # 顶层（无缩进）
        维度A:        # 二级（两个空格缩进）
          - 选项1     # 三级（四个空格缩进，连字符列表）
          - 选项2
        维度B:
          - ...
    - 返回 { 维度: [选项1, 选项2, ...] } 的有序字典（按出现顺序）。
    """
    from collections import OrderedDict
    result = OrderedDict()
    in_specs = False
    cur_key = None
    try:
        with result_path.open("r", encoding="utf-8") as f:
            for raw in f:
                line = raw.rstrip("\n\r")
                if not line.strip():
                    continue
                # 顶层注释跳过
                if not in_specs and line.lstrip().startswith("#"):
                    continue
                # 进入 specs 顶层
                if not in_specs and line.strip() == "specs:":
                    in_specs = True
                    cur_key = None
                    continue
                # 一旦进入 specs，遇到新的顶层键（非空且不以空格开头且以冒号结尾）则退出
                if in_specs and (not line.startswith(" ")) and line.strip().endswith(":"):
                    break
                if not in_specs:
                    continue
                # 二级键：两个空格起始，且以冒号结尾
                if line.startswith("  ") and (not line.startswith("   ")) and line.strip().endswith(":"):
                    cur_key = line.strip()[:-1].strip()
                    if cur_key:
                        result[cur_key] = []
                    continue
                # 三级子项：四个空格起始，且包含连字符
                if cur_key and line.startswith("    ") and line.lstrip().startswith("-"):
                    dash_idx = line.find("-")
                    val = line[dash_idx + 1 :].strip()
                    # 去掉包裹引号
                    if (val.startswith('"') and val.endswith('"')) or (val.startswith("'") and val.endswith("'")):
                        val = val[1:-1]
                    if val:
                        result[cur_key].append(val)
        if not result:
            raise RuntimeError("未能从 result.yml 的 specs 段解析到任何维度/选项，请检查文件结构是否符合预期")
    except Exception as e:
        raise RuntimeError(f"解析最新 result.yml 失败：{e}")
    print(f"[信息] 从 result.yml 解析到的规格维度: {list(result.keys())}")
    return result


def read_spec_dimensions_from_latest_output() -> dict:
    """读取 output/ 下最新子目录的 result.yml，并返回 {维度: [选项...]}。

    若未找到任何 result.yml，将给出中文错误提示。
    """
    p = _find_latest_result_yml()
    if not p:
        raise RuntimeError("未在 output/ 下找到任何 result.yml，请先生成规格数据后再运行本任务")
    print(f"[信息] 使用规格配置文件: {p}")
    return _parse_specs_from_result_yml(p)




# -------------------------
# 价格计划：从 result.yml 读取 dims 与 combos，生成按表格顺序的价格列表
# -------------------------
def _parse_dims_and_prices_from_result_yml(result_path: Path) -> Tuple[List[str], List[Optional[float]]]:
    """最简 YAML 解析 dims 与 combos：
    - 仅识别如下层级结构：
      dims:
        - 维度A
        - 维度B
      combos:
        - [iA, iB, ..., price, img_idx]
    - 返回 (dims列表, price列表)。其中 price 列表顺序为 combos 出现顺序；
      price 可能为 None（若 result.yml 中为 null）。
    """
    dims: List[str] = []
    prices: List[Optional[float]] = []
    in_dims = False
    in_combos = False
    try:
        import ast
        with result_path.open("r", encoding="utf-8") as f:
            for raw in f:
                line = raw.rstrip("\n\r")
                if not line.strip():
                    continue
                # 注释行直接跳过
                if line.lstrip().startswith("#"):
                    continue
                # 顶层键切换
                if not line.startswith(" ") and line.strip().endswith(":"):
                    key = line.strip()[:-1].strip()
                    in_dims = (key == "dims")
                    in_combos = (key == "combos")
                    # 进入新段时继续读取下一行
                    continue
                # 读取 dims
                if in_dims:
                    # 只接受形如 "  - 项" 的列表项
                    if line.startswith("  ") and line.lstrip().startswith("-"):
                        dash_idx = line.find("-")
                        val = line[dash_idx + 1 :].strip()
                        # 去掉包裹引号
                        if (val.startswith('"') and val.endswith('"')) or (val.startswith("'") and val.endswith("'")):
                            val = val[1:-1]
                        if val:
                            dims.append(val)
                        continue
                    # 碰到非列表行或缩进不匹配，视为 dims 结束
                    in_dims = False
                # 读取 combos
                if in_combos:
                    ls = line.lstrip()
                    if ls.startswith("-") and "[" in ls and "]" in ls:
                        # 提取方括号内容
                        try:
                            content = ls[ls.find("[") : ls.rfind("]") + 1]
                            # 兼容 yaml 的 null
                            content = content.replace("null", "None")
                            arr = ast.literal_eval(content)
                            if isinstance(arr, (list, tuple)) and len(arr) >= len(dims) + 1:
                                price = arr[len(dims)]
                                # 允许 None / 数值
                                if price is None:
                                    prices.append(None)
                                else:
                                    try:
                                        prices.append(float(price))
                                    except Exception:
                                        prices.append(None)
                            else:
                                prices.append(None)
                        except Exception:
                            prices.append(None)
                        continue
                    # 碰到不是以 "- [" 开头的行，认为 combos 结束
                    if not (ls.startswith("-") and ("[" in ls)):
                        in_combos = False
        if not dims:
            raise RuntimeError("未能从 result.yml 的 dims 段解析到任何维度，请检查文件结构是否符合预期")
        if not prices:
            # 允许没有 combos（例如仅有维度与选项），此时返回空价格计划
            print("[提示] result.yml 中未发现 combos 段或未解析到任何价格，价格填充将被跳过")
        return dims, prices
    except Exception as e:
        raise RuntimeError(f"解析最新 result.yml 的 dims/combos 失败：{e}")


def read_price_plan_from_latest_output() -> List[Optional[float]]:
    """读取 output/ 下最新子目录的 result.yml，返回按 combos 顺序的价格列表。

    - 列表顺序即为表格的渲染顺序（与样例一致，按 dims 组合的嵌套顺序）。
    - 若解析失败或无 combos，将返回空列表。
    """
    p = _find_latest_result_yml()
    if not p:
        print("[提示] 未在 output/ 下找到任何 result.yml，价格填充将被跳过")
        return []
    print(f"[信息] 使用规格配置文件(价格)：{p}")
    _, prices = _parse_dims_and_prices_from_result_yml(p)
    print(f"[信息] 已解析到 {len(prices)} 条价格计划")
    return prices


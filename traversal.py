import time
import random
import itertools
from typing import List, Tuple
from selenium.webdriver.common.by import By

from common import log_debug, normalize_price_text, append_to_log
from sku_utils import (
    SkuOption,
    SkuDimension,
    SKU_ITEM_SELECTOR,
    SKU_OPTION_SELECTOR,
    get_price_text,
    get_main_image_url,
    is_selected_element,
    read_current_selected_vids,
)


def generate_all_combinations(sku_dimensions: List[SkuDimension]) -> List[Tuple[SkuOption, ...]]:
    """生成所有SKU选项的笛卡尔积组合。"""
    return list(itertools.product(*[d.options for d in sku_dimensions]))


essential_float_delay = (0.02, 0.06)


def ensure_combination_selected(driver, combination: List[SkuOption], last_selected_vids: List[str] | None = None) -> List[str]:
    """按需点击组合中的 SKU 选项：仅对发生变化的维度执行点击；随后做一次快速校验，必要时补点。
    返回本次目标组合的 vid 列表，供下次迭代复用，减少无效点击。
    """
    need_change_indices = list(range(len(combination)))
    if last_selected_vids and len(last_selected_vids) == len(combination):
        need_change_indices = [i for i, opt in enumerate(combination) if last_selected_vids[i] != opt.vid]

    if not need_change_indices:
        log_debug(f"点击SKU: 本次无需变更（沿用上次选择），维度索引 {list(range(len(combination)))}")
        return [opt.vid for opt in combination]

    # 首轮：只点击有变化的维度
    for dim_idx in need_change_indices:
        option = combination[dim_idx]
        try:
            sku_items = driver.find_elements(By.CSS_SELECTOR, SKU_ITEM_SELECTOR)
            current_dim = sku_items[dim_idx]
            option_selector = f'{SKU_OPTION_SELECTOR}[data-vid="{option.vid}"]'
            option_element = current_dim.find_element(By.CSS_SELECTOR, option_selector)
            if not is_selected_element(option_element):
                driver.execute_script("arguments[0].click();", option_element)
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
            if not is_selected_element(option_element):
                driver.execute_script("arguments[0].click();", option_element)
                time.sleep(0.03)
    except Exception:
        pass

    return [opt.vid for opt in combination]


def reorder_with_current_selected(driver, sku_dimensions: List[SkuDimension], combinations: List[Tuple[SkuOption, ...]]):
    """将当前页面已选中的组合提前到第一个，并返回初始化的 last_selected_vids。"""
    last_selected_vids: List[str] = []
    try:
        current_vids = read_current_selected_vids(driver, len(sku_dimensions))
        log_debug(f"当前已选VID: {current_vids}")
        if current_vids and len(current_vids) == len(sku_dimensions):
            cur_opts: List[SkuOption] = []
            for i, dim in enumerate(sku_dimensions):
                vid = current_vids[i]
                if not vid:
                    cur_opts = []
                    break
                found = next((o for o in dim.options if o.vid == vid), None)
                if not found:
                    cur_opts = []
                    break
                cur_opts.append(found)
            if cur_opts:
                cur_tuple = tuple(cur_opts)
                if cur_tuple in combinations:
                    combinations = combinations[:]
                    combinations.remove(cur_tuple)
                    combinations.insert(0, cur_tuple)
                    log_debug("已将当前已选组合置于遍历首位")
                    last_selected_vids = current_vids[:]
    except Exception as e:
        log_debug(f"处理当前已选组合时异常: {e}")
    return combinations, last_selected_vids


def apply_max_combos_limit(combinations: List[Tuple[SkuOption, ...]]) -> List[Tuple[SkuOption, ...]]:
    """根据环境变量 MAX_COMBOS 限制前 N 个组合用于快速自测。"""
    import os
    try:
        max_combos_str = os.environ.get("MAX_COMBOS", "").strip()
        if max_combos_str:
            max_combos = int(max_combos_str)
            if max_combos > 0 and max_combos < len(combinations):
                print(f"[信息] 仅执行前 {max_combos} 个组合用于快速验证（通过 MAX_COMBOS 控制）")
                return combinations[:max_combos]
    except Exception:
        pass
    return combinations


def handle_single_combination(
    driver,
    sku_dimensions: List[SkuDimension],
    combination: Tuple[SkuOption, ...],
    last_selected_vids: List[str] | None,
    t_after_parse: float,
    combo_idx: int,
    total_count: int,
    first_select_logged: bool,
):
    """处理单个组合：点击、取价、日志记录。返回结果行（或None）与更新后的 last_selected_vids。"""
    try:
        print(f"\n[进度] 处理组合 {combo_idx}/{total_count}")
        combo_text_parts = [f"{sku_dimensions[i].name}: {opt.text}" for i, opt in enumerate(combination)]
        print(f"[组合] {' | '.join(combo_text_parts)}")

        combo_info = "---".join([opt.text for opt in combination])

        start_time = time.time()
        t_click_begin = time.perf_counter()
        last_selected_vids = ensure_combination_selected(
            driver, list(combination), last_selected_vids or None
        )
        t_click_end = time.perf_counter()
        price = get_price_text(driver)
        image_url = get_main_image_url(driver)
        t_price_end = time.perf_counter()
        # 调试输出主图区域图片URL（通常为规格图），便于快速定位问题
        try:
            print(f"[图片] 链接: {image_url if image_url else '空'}")
            try:
                append_to_log(f"图片链接: {image_url if image_url else '空'}")
            except Exception:
                pass
        except Exception:
            pass
        if not first_select_logged:
            first_select_logged = True
            log_debug(
                f"首次选中耗时：点击 {(t_click_end - t_click_begin)*1000:.0f}ms；取价 {(t_price_end - t_click_end)*1000:.0f}ms；自页面就绪起 {(t_click_end - t_after_parse):.3f}s"
            )

        price = normalize_price_text(price)

        # 结果行：各维度 + 图片(用于嵌入) + 图片链接(纯文本；主图区域展示的规格图) + 价格
        result_row = [opt.text for opt in combination] + [image_url, image_url, price]
        print(f"[成功] 价格: {price}")
        elapsed_time = time.time() - start_time
        print(f"[耗时] 本组合用时 {elapsed_time:.3f} 秒")

        full_combo_info = f"{combo_info}---{price}"
        append_to_log(f"完成组合 {combo_idx}: {full_combo_info}")

        time.sleep(random.uniform(*essential_float_delay))

        return result_row, last_selected_vids, first_select_logged
    except Exception as e:
        print(f"[错误] 处理组合 {combo_idx} 失败: {e}")
        return None, last_selected_vids or [], first_select_logged


def traverse_and_collect(
    driver,
    sku_dimensions: List[SkuDimension],
    combinations: List[Tuple[SkuOption, ...]],
    last_selected_vids: List[str],
    t_after_parse: float,
):
    """遍历所有组合并汇总结果。返回结果表与成功计数。"""
    results: List[List[str]] = []
    success_count = 0
    first_select_logged = False
    for combo_idx, combination in enumerate(combinations, 1):
        result_row, last_selected_vids, first_select_logged = handle_single_combination(
            driver,
            sku_dimensions,
            combination,
            last_selected_vids,
            t_after_parse,
            combo_idx,
            len(combinations),
            first_select_logged,
        )
        if result_row is not None:
            results.append(result_row)
            success_count += 1
    return results, success_count

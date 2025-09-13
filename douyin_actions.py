# -*- coding: utf-8 -*-
"""
抖店页面元素与动作封装：

- 统一定义与 "元素示例/" 对应的选择器
- 创建维度、打开下拉、在下拉中选择/创建维度
- 在维度容器内批量输入所有选项（极速模式）
- 点击“保存草稿”、收集已存在维度

说明：严格遵循项目已有选择器与中文日志规范；避免扩大需求，仅覆盖当前任务需要的最小集合。
"""
from __future__ import annotations

import time
from typing import List, Set

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, WebDriverException


# -------------------------
# 元素选择器（严格参考 元素示例/ 下的文件）
# -------------------------
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
# - “请输入规格类型” 输入框
X_INPUT_CREATE_TYPE = (
    By.XPATH,
    "//input[contains(@placeholder,'请输入规格类型')]",
)
# - 页面中已添加的维度展示
X_EXISTING_SPEC_ITEMS = (
    By.XPATH,
    "//span[contains(@class,'ecom-g-select-selection-item')]",
)

# - 规格值输入容器与输入框（元素示例/商品规格区域.html）
SPEC_VALUE_CONTAINER_XPATH_TMPL = "//div[@id=concat('skuValue-', '{dim}')]"
SPEC_VALUE_INPUT_XPATH_TMPL = ".//input[contains(@placeholder,'请输入') and contains(@placeholder, '{dim}') and @type='text']"

# - 保存草稿按钮（元素示例/保存草稿按钮.html）
X_SAVE_DRAFT_BUTTON = (
    By.XPATH,
    "//button[.//span[normalize-space(text())='保存草稿']]",
)


def click_add_spec_button(driver) -> None:
    # 记录点击前可用的“请选择规格类型”下拉触发器数量
    before_cnt = 0
    try:
        before_cnt = len(driver.find_elements(*X_SPEC_DROPDOWN_TRIGGER))
    except Exception:
        before_cnt = 0

    btn = WebDriverWait(driver, 20).until(EC.element_to_be_clickable(X_ADD_SPEC_BUTTON))
    try:
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
    except Exception:
        pass
    try:
        btn.click()
    except Exception:
        try:
            driver.execute_script("arguments[0].click();", btn)
        except Exception:
            raise
    # 确认新增的占位“请选择规格类型”出现，防止后续找不到下拉而卡住
    try:
        WebDriverWait(driver, 10).until(lambda d: len(d.find_elements(*X_SPEC_DROPDOWN_TRIGGER)) > before_cnt)
    except Exception:
        # 兜底再滚动一次页面底部重试一次探测
        try:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        except Exception:
            pass
        WebDriverWait(driver, 5).until(lambda d: len(d.find_elements(*X_SPEC_DROPDOWN_TRIGGER)) > before_cnt)
    print("[步骤] 已点击“添加规格类型”按钮，并检测到新的下拉触发器出现")


def open_latest_spec_dropdown(driver) -> None:
    nodes = WebDriverWait(driver, 20).until(EC.presence_of_all_elements_located(X_SPEC_DROPDOWN_TRIGGER))
    last = nodes[-1]
    try:
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", last)
    except Exception:
        pass
    try:
        last.click()
    except Exception:
        driver.execute_script("arguments[0].click();", last)
    print("[步骤] 已点击“规格类型下拉按钮”（占位：请选择规格类型）")


def select_dimension_from_dropdown(driver, dim: str) -> bool:
    dropdown = WebDriverWait(driver, 20).until(EC.presence_of_element_located(X_VISIBLE_DROPDOWN))
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


def create_dimension_via_dialog(driver, dim: str) -> None:
    dropdown = WebDriverWait(driver, 10).until(EC.presence_of_element_located(X_VISIBLE_DROPDOWN))
    try:
        create_link = dropdown.find_element(*X_CREATE_TYPE_LINK)
        create_link.click()
        print("[步骤] 已点击“创建类型”")
    except Exception:
        print("[错误] 未在下拉中找到“创建类型”链接，无法创建新维度")
        raise
    try:
        input_box = WebDriverWait(driver, 10).until(EC.presence_of_element_located(X_INPUT_CREATE_TYPE))
        input_box.clear()
        input_box.send_keys(dim)
        input_box.send_keys(Keys.ENTER)
        print(f"[步骤] 已输入新维度并回车：{dim}")
    except Exception:
        print("[错误] 未找到“请输入规格类型”的输入框，创建失败")
        raise
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, f"//span[contains(@class,'ecom-g-select-selection-item') and normalize-space(text())='{dim}']"))
    )


def click_save_draft(driver) -> None:
    try:
        btn = WebDriverWait(driver, 20).until(EC.element_to_be_clickable(X_SAVE_DRAFT_BUTTON))
        btn.click()
        print("[步骤] 已点击“保存草稿”按钮")
    except Exception as e:
        print(f"[错误] 未能点击“保存草稿”按钮：{e}")


def collect_existing_dimensions(driver) -> Set[str]:
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


def _get_spec_container(driver, dim: str):
    container_xpath = SPEC_VALUE_CONTAINER_XPATH_TMPL.format(dim=dim)
    try:
        return WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, container_xpath))
        )
    except Exception:
        return None


def _option_exists_in_dimension(driver, dim: str, val: str) -> bool:
    container = _get_spec_container(driver, dim)
    if not container:
        return False
    try:
        nodes_text = container.find_elements(By.XPATH, f".//*[normalize-space(text())='{val}']")
        if len(nodes_text) > 0:
            return True
        nodes_input = container.find_elements(By.XPATH, f".//input[@type='text' and @value='{val}']")
        if len(nodes_input) > 0:
            return True
        return False
    except Exception:
        return False


def _list_existing_options_set(driver, dim: str) -> Set[str]:
    s: Set[str] = set()
    container = _get_spec_container(driver, dim)
    if not container:
        return s
    try:
        nodes_inp = container.find_elements(By.XPATH, ".//input[@type='text']")
        for n in nodes_inp:
            v = (n.get_attribute("value") or "").strip()
            if v:
                s.add(v)
    except Exception:
        pass
    try:
        nodes_txt = container.find_elements(
            By.XPATH,
            ".//div[contains(@class,'style_skuValue__') or contains(@class,'style_skuValue__OsgA1')]//*[normalize-space(text())!='']",
        )
        for n in nodes_txt:
            t = (n.text or "").strip()
            if t:
                s.add(t)
    except Exception:
        pass
    return s


def _find_spec_value_input(driver, dim: str):
    container_xpath = SPEC_VALUE_CONTAINER_XPATH_TMPL.format(dim=dim)
    try:
        container = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, container_xpath))
        )
        try:
            return container.find_element(
                By.XPATH,
                ".//div[contains(@class,'style_skuValueInput__oQFaa') and contains(@class,'style_forCreate__tNN3f')]//input[contains(@placeholder,'请输入') and contains(@placeholder, '%s') and @type='text' and (not(@value) or @value='')]" % dim,
            )
        except Exception:
            pass
        try:
            return container.find_element(
                By.XPATH,
                ".//input[contains(@placeholder,'请输入') and contains(@placeholder, '%s') and @type='text']" % dim,
            )
        except Exception:
            pass
    except Exception:
        pass
    try:
        return WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, f"//input[contains(@placeholder,'{dim}') and contains(@placeholder,'请输入') and @type='text' and (not(@value) or @value='')]"))
        )
    except Exception:
        try:
            return WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, f"//input[contains(@placeholder,'{dim}') and contains(@placeholder,'请输入') and @type='text']"))
            )
        except Exception:
            return None


def input_options_for_dimension(driver, dim: str, options: List[str]) -> None:
    if not options:
        print(f"[信息] 维度“{dim}”未配置任何选项，跳过输入")
        return

    existing_set = _list_existing_options_set(driver, dim)
    pending = [v for v in options if v not in existing_set]
    if not pending:
        print(f"[信息] 维度“{dim}”的所有选项均已存在，跳过输入")
        return

    def _fast_type_one(val: str) -> bool:
        for attempt in range(2):
            try:
                inp = _find_spec_value_input(driver, dim)
                if not inp:
                    return False
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", inp)
                inp.send_keys(val)
                inp.send_keys(Keys.ENTER)
                time.sleep(0.01)
                return True
            except (StaleElementReferenceException, WebDriverException):
                time.sleep(0.01)
                continue
            except Exception:
                return False
        return False

    try:
        driver.implicitly_wait(0)
        for v in list(pending):
            if _option_exists_in_dimension(driver, dim, v):
                pending.remove(v)
                continue
            ok = _fast_type_one(v)
            if not ok:
                print(f"[警告] 快速输入失败（维度={dim}，选项={v}），稍后重试")

        existing_set2 = _list_existing_options_set(driver, dim)
        missing = [v for v in pending if v not in existing_set2]
        if not missing:
            return
        for v in missing:
            ok = _fast_type_one(v)
            if not ok:
                print(f"[警告] 仍未能添加该选项（维度={dim}，选项={v}），请稍后手动检查或重跑任务")
    finally:
        try:
            driver.implicitly_wait(2)
        except Exception:
            pass


# -------------------------
# SKU 价格录入（参考：元素示例/sku表格区域.html, 元素示例/sku价格输入框.html）
# -------------------------
# 虚拟表格主体容器
X_SKU_TABLE_VIRTUAL_BODY = (
    By.XPATH,
    "//div[contains(@class,'ecom-g-table-tbody-virtual') and contains(@class,'ecom-g-table-tbody')]",
)


def _find_price_inputs(driver):
    """查找当前可见的所有 SKU 价格输入框（不滚动）。
    限定范围：虚拟表格体内、非 extra 行、价格列中的 ecom-g-input-number-input。
    """
    try:
        body = WebDriverWait(driver, 10).until(EC.presence_of_element_located(X_SKU_TABLE_VIRTUAL_BODY))
    except Exception:
        return []
    try:
        return body.find_elements(
            By.XPATH,
            ".//tr[contains(@class,'ecom-g-table-row') and not(contains(@class,'ecom-g-table-row-extra'))]//td[contains(@class,'attr-column-field_price')]//input[contains(@class,'ecom-g-input-number-input')]",
        )
    except Exception:
        return []


def _normalize_price_text(price: float) -> str:
    """将数值价格格式化为输入框可接受的字符串，且满足最小 0.01 的校验。"""
    try:
        if price is None:
            return ""
        if price < 0.01:
            return ""
        if abs(price - int(price)) < 1e-9:
            return str(int(price))
        txt = f"{price:.2f}"
        if "." in txt:
            txt = txt.rstrip("0").rstrip(".")
        return txt
    except Exception:
        return ""


def fill_prices_for_sku_table(driver, prices: List[float]) -> None:
    """为 SKU 表格的价格列依次填写价格：
    - 严格顺序：第 N 个输入框对应第 N 条价格；
    - 幂等：若输入框已有 >0 的值，则跳过不覆盖；
    - 价格为空或 <0.01 时跳过；
    - 表格为虚拟滚动：逐批填写并滚动加载下一批。
    """
    if not prices:
        print("[提示] 未解析到任何价格数据（或 combos 为空），跳过价格填充")
        return
    try:
        body = WebDriverWait(driver, 20).until(EC.presence_of_element_located(X_SKU_TABLE_VIRTUAL_BODY))
    except Exception:
        print("[提示] 未找到 SKU 表格区域，价格填充将被跳过")
        return

    # 将页面滚动到顶部，尽量与表格顺序对齐
    try:
        driver.execute_script("window.scrollTo(0, 0);")
    except Exception:
        pass

    idx = 0  # 已处理到的价格计划索引
    stagnant_rounds = 0
    MAX_STAGNANT_ROUNDS = 8

    while idx < len(prices):
        inputs = _find_price_inputs(driver)
        if not inputs:
            stagnant_rounds += 1
            if stagnant_rounds >= MAX_STAGNANT_ROUNDS:
                break
            time.sleep(0.2)
            continue

        stagnant_rounds = 0
        for inp in inputs:
            if idx >= len(prices):
                break
            price = prices[idx]
            idx += 1
            if price is None:
                continue
            try:
                cur_val = (inp.get_attribute("value") or "").strip()
                try:
                    cur_num = float(cur_val) if cur_val else 0.0
                except Exception:
                    cur_num = 0.0
                # 已有值且>0，视为已填，跳过
                if cur_num > 0:
                    continue

                txt = _normalize_price_text(float(price))
                if not txt:
                    continue

                try:
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", inp)
                except Exception:
                    pass
                try:
                    inp.click()
                except Exception:
                    try:
                        driver.execute_script("arguments[0].click();", inp)
                    except Exception:
                        pass
                try:
                    inp.send_keys(Keys.CONTROL, "a")
                except Exception:
                    try:
                        inp.clear()
                    except Exception:
                        pass
                inp.send_keys(txt)
                time.sleep(0.02)
                try:
                    inp.send_keys(Keys.ENTER)
                except Exception:
                    pass
            except StaleElementReferenceException:
                continue
            except Exception as e:
                print(f"[警告] 写入价格失败（索引={idx-1}，价格={price}）：{e}")

        # 滚动以触发后续行渲染
        try:
            inputs2 = _find_price_inputs(driver)
            if inputs2:
                try:
                    driver.execute_script("arguments[0].scrollIntoView({block:'end'});", inputs2[-1])
                except Exception:
                    try:
                        driver.execute_script("arguments[0].scrollTop = arguments[0].scrollTop + 600;", body)
                    except Exception:
                        pass
        except Exception:
            pass
        time.sleep(0.1)

    print(f"[信息] 价格填充完成（计划条数={len(prices)}，已处理到索引={idx}）")


# -*- coding: utf-8 -*-
"""
抖店-创建规格维度 RPA（任务编排入口）

说明：
- 从 conf/douyin/product-url.txt 读取商品草稿链接（若不存在则回退 conf/product-url.txt）
- 从 conf/规格.yml 读取需要创建的“维度”与每个维度的“选项”（顶层键名为维度，缩进的“- 项”为选项）
- 打开浏览器访问草稿页，依次点击“添加规格类型”，展开“规格类型下拉按钮”，在下拉列表中选择维度；
  若列表没有该维度，则点击“创建类型”，在弹出的输入框中输入维度名并回车。
- 代码内全部中文日志与注释，严格参考 “元素示例/” 下的文件来定位元素；录入完所有维度选项后点击“保存草稿”。

注意：
- 为减少耦合，本文件仅保留任务编排，通用配置与浏览器/页面操作已拆分到独立模块。
- 若页面需登录，请先手动完成登录后再继续。
"""
from __future__ import annotations

import time

from robocorp.tasks import task

from config import SPECS_YAML_FILE, read_product_url, read_spec_dimensions_with_options
from driver_utils import build_edge_driver, navigate_to_url, wait_for_login_and_page_ready
from douyin_actions import (
    collect_existing_dimensions,
    click_add_spec_button,
    open_latest_spec_dropdown,
    select_dimension_from_dropdown,
    create_dimension_via_dialog,
    input_options_for_dimension,
    click_save_draft,
)


@task
def create_douyin_spec_dimensions() -> None:
    """创建抖店商品规格维度，并录入每个维度的所有选项，最后保存草稿。"""
    print("[开始] 抖店-创建规格维度 任务启动…")
    url = read_product_url()
    dims_map = read_spec_dimensions_with_options(SPECS_YAML_FILE)
    dims = list(dims_map.keys())

    driver = build_edge_driver()
    try:
        print(f"[步骤] 打开商品草稿链接: {url}")
        navigate_to_url(driver, url)
        wait_for_login_and_page_ready(driver, timeout=60)

        existed = collect_existing_dimensions(driver)

        for dim in dims:
            if dim not in existed:
                # 先创建维度
                click_add_spec_button(driver)
                open_latest_spec_dropdown(driver)
                if not select_dimension_from_dropdown(driver, dim):
                    create_dimension_via_dialog(driver, dim)
                # 更新已存在集合
                time.sleep(0.5)
                existed = collect_existing_dimensions(driver)
            else:
                print(f"[信息] 维度已存在：{dim}，将直接录入选项")
            # 无论此前是否存在，均录入该维度的所有选项
            try:
                input_options_for_dimension(driver, dim, dims_map.get(dim, []))
            except Exception as e:
                print(f"[警告] 录入维度选项发生异常（维度={dim}）：{e}")

        # 保存草稿
        click_save_draft(driver)
        print("[完成] 规格维度与选项录入并已尝试保存草稿。")
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




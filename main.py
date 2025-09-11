from robocorp.tasks import task
import time
from pathlib import Path
import re

# 模块化导入
from common import read_product_url, log_debug
from browser_utils import (
    prepare_clean_edge_state,
    init_edge_driver,
    open_product_page,
    kill_driver_processes,
)
from sku_utils import (
    parse_sku_dimensions,
    print_dimensions_summary,
    write_dimensions_structure_log,
    get_product_name,
    get_shop_name,
)
from traversal import (
    generate_all_combinations,
    reorder_with_current_selected,
    apply_max_combos_limit,
    traverse_and_collect,
)
from io_utils import export_results_to_excel


@task
def traverse_all_sku_combinations():
    """
    自动遍历所有SKU维度组合，获取每个组合的价格信息，并写入Excel表格。
    使用防检测机制：随机延迟、模拟人类行为、控制点击速度。
    """
    url = read_product_url()
    print(f"[步骤] 读取到商品链接: {url}")

    # 准备浏览器环境并初始化驱动
    t_task0 = time.perf_counter()
    prepare_clean_edge_state()
    t_after_clean = time.perf_counter()
    driver = init_edge_driver()
    t_after_init = time.perf_counter()
    log_debug(f"阶段耗时：清理 {(t_after_clean - t_task0):.3f}s；初始化驱动 {(t_after_init - t_after_clean):.3f}s")

    try:
        # 打开页面并等待初始加载
        open_product_page(driver, url, wait_timeout=30)
        t_after_open = time.perf_counter()

        # 解析维度
        sku_dimensions = parse_sku_dimensions(driver)
        t_after_parse = time.perf_counter()
        log_debug(f"阶段耗时：打开页面 {(t_after_open - t_after_init):.3f}s；解析SKU {(t_after_parse - t_after_open):.3f}s")

        # 解析店铺名与商品名（用于导出目录名）
        shop_name = get_shop_name(driver)
        product_name = get_product_name(driver)
        print(f"[信息] 店铺: {shop_name} | 商品: {product_name}")

        # 概要信息与结构日志
        print_dimensions_summary(sku_dimensions)
        write_dimensions_structure_log(sku_dimensions)

        # 组合生成与预处理
        combinations = generate_all_combinations(sku_dimensions)
        combinations, last_selected_vids = reorder_with_current_selected(driver, sku_dimensions, combinations)
        combinations = apply_max_combos_limit(combinations)

        print("[步骤] 开始遍历所有SKU组合...")

        # 遍历并收集
        results, success_count = traverse_and_collect(
            driver=driver,
            sku_dimensions=sku_dimensions,
            combinations=combinations,
            last_selected_vids=last_selected_vids if 'last_selected_vids' in locals() else [],
            t_after_parse=t_after_parse,
        )

        print(f"\n[步骤] 遍历完成！成功处理 {success_count}/{len(combinations)} 个组合")

        # 导出结果到 Excel（包含表头）：各维度 + 图片 + 图片链接 + 价格
        headers = [dim.name for dim in sku_dimensions] + ["图片", "图片链接", "价格"]

        # 根据“店铺名+商品名”建立子目录
        def _safe_name(n: str) -> str:
            try:
                s = str(n)
            except Exception:
                s = ""
            s = re.sub(r"[\\/:*?\"<>|]+", " ", s)
            s = re.sub(r"\s+", " ", s).strip()
            return (s[:80] if s else "未命名")

        folder_name = f"[{_safe_name(shop_name)}]{_safe_name(product_name)}"
        base_output = Path(__file__).resolve().parent / "output"
        output_dir = base_output / folder_name
        output_dir.mkdir(parents=True, exist_ok=True)
        excel_path = output_dir / "result.xlsx"
        try:
            export_results_to_excel(results, headers, excel_path)
            print(f"[步骤] 已导出结果到: {excel_path}")
        except Exception as e:
            print(f"[错误] 导出Excel失败: {e}")

    finally:
        try:
            driver.quit()
        except Exception:
            pass


def main():
    """
    RPA任务统一入口函数：按顺序执行所有RPA任务。
    """
    try:
        print("=" * 50)
        print("开始执行RPA任务...")
        print("=" * 50)

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
        kill_driver_processes()

    return 0


if __name__ == "__main__":
    import sys
    sys.exit(main())

# -*- coding: utf-8 -*-
"""
配置与路径常量、基础数据读取函数

- 路径常量：项目根目录、配置目录、驱动路径等
- 读取函数：浏览器路径、商品链接、规格配置（最简 YAML 解析）

说明：严格保持中文日志与报错，遵循项目原有规范；尽量少依赖第三方库。
"""
from __future__ import annotations

from pathlib import Path
from typing import Optional


# -------------------------
# 路径与常量
# -------------------------
PROJECT_ROOT = Path(__file__).resolve().parent
CONF_DIR = PROJECT_ROOT / "conf"
DRIVER_PATH = PROJECT_ROOT / "driver" / "msedgedriver.exe"
BROWSER_PATH_FILE = CONF_DIR / "browser.txt"
TEMP_DIR = PROJECT_ROOT / "temp"

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




from dataclasses import dataclass
from typing import List, Tuple
import time
import random
from pathlib import Path
from selenium.webdriver.common.by import By

from common import log_debug, normalize_price_text

# 统一的页面元素选择器常量，便于维护与复用（避免写死哈希后缀）
SKU_ITEM_SELECTOR = "[class*='skuItem']"
SKU_OPTION_SELECTOR = "[class*='valueItem']:not([class*='ImgWrap'])"
DIM_LABEL_SELECTOR = "[class*='ItemLabel'] span.f-els-2"
OPTION_TEXT_SELECTOR = "span.f-els-1"

# 价格相关选择器（主选择器仍优先，配合兜底）
PRICE_MAIN_TEXT = "[class*='highlightPrice'] [class*='text']"
PRICE_SYMBOL = "[class*='highlightPrice'] [class*='symbol']"
ORIG_PRICE_TEXTS = "[class*='subPrice'] [class*='text']"
PRICE_ALT_SELECTORS = [
    "[class*='beltPrice'] [class*='text']",
    "[class*='priceWrap'] [class*='text']",
    "[class*='price'] [class*='number']",
    ".tm-price-current",
    "[class*='highlightPrice'] [class*='text']",
]

# 主图相关选择器（方案A）
MAIN_PIC_IMG_SELECTOR = "img[class*='mainPic']"
ZOOM_IMG_DIV_SELECTOR = ".js-image-zoom__zoomed-image"


# 数据模型：更通用、更可读
@dataclass(frozen=True)
class SkuOption:
    vid: str
    text: str


@dataclass
class SkuDimension:
    name: str
    options: List[SkuOption]


def parse_sku_dimensions(driver) -> List[SkuDimension]:
    """解析页面上的所有 SKU 维度与选项。"""
    print("[步骤] 解析SKU维度和选项...")
    t0 = time.perf_counter()
    js = (
        "return (function(itemSel, labelSel, optionSel, textSel){\n"
        "  var items = Array.from(document.querySelectorAll(itemSel));\n"
        "  var dims = [];\n"
        "  for (var i=0; i<items.length; i++){\n"
        "    var item = items[i];\n"
        "    var label = item.querySelector(labelSel);\n"
        "    var name = label ? (label.getAttribute('title') || label.textContent || ('维度' + (i+1))) : ('维度' + (i+1));\n"
        "    var opts = [];\n"
        "    var nodes = Array.from(item.querySelectorAll(optionSel));\n"
        "    for (var j=0; j<nodes.length; j++){\n"
        "      var el = nodes[j];\n"
        "      var dis = (el.getAttribute('data-disabled') || '').toLowerCase() === 'true';\n"
        "      if (dis) continue;\n"
        "      var vid = el.getAttribute('data-vid');\n"
        "      var span = el.querySelector(textSel);\n"
        "      var txt = span ? (span.getAttribute('title') || span.textContent || '').trim() : '';\n"
        "      if (vid && txt){ opts.push({vid: vid, text: txt}); }\n"
        "    }\n"
        "    if (opts.length){ dims.push({name: name.trim(), options: opts}); }\n"
        "  }\n"
        "  return dims;\n"
        "})(arguments[0], arguments[1], arguments[2], arguments[3]);"
    )

    dims_data = driver.execute_script(
        js,
        SKU_ITEM_SELECTOR,
        DIM_LABEL_SELECTOR,
        f"{SKU_OPTION_SELECTOR}[data-vid]",
        OPTION_TEXT_SELECTOR,
    ) or []

    sku_dimensions: List[SkuDimension] = []
    try:
        for d in dims_data:
            name = (d.get('name') or '').strip() or '未命名维度'
            opts = [SkuOption(vid=str(o.get('vid') or ''), text=str(o.get('text') or '').strip()) for o in (d.get('options') or [])]
            if opts:
                sku_dimensions.append(SkuDimension(name=name, options=opts))
    except Exception as e:
        log_debug(f"JS 解析SKU异常: {e}")

    log_debug(f"解析SKU维度完成：{len(sku_dimensions)} 个维度；耗时 {(time.perf_counter() - t0):.3f}s")
    return sku_dimensions


def is_selected_element(elem) -> bool:
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


def read_current_selected_vids(driver, dims_count: int) -> List[str]:
    """读取当前页面中每个维度已选中的 data-vid（若未选中返回空字符串）。"""
    vids: List[str] = []
    try:
        sku_items = driver.find_elements(By.CSS_SELECTOR, SKU_ITEM_SELECTOR)
        for idx in range(min(dims_count, len(sku_items))):
            item = sku_items[idx]
            selected = ""
            try:
                candidates = item.find_elements(By.CSS_SELECTOR, f"{SKU_OPTION_SELECTOR}[data-vid]")
                for el in candidates:
                    if is_selected_element(el):
                        v = el.get_attribute("data-vid") or ""
                        if v:
                            selected = v
                            break
            except Exception:
                pass
            vids.append(selected)
    except Exception:
        pass
    return vids


def get_price_text(driver) -> str:
    """获取当前所选组合的价格文本（主选择器优先 + 包含匹配兜底；JS一次性查询 + 短轮询）。"""
    js = (
        "return (function(mainSel, symSel, alts, beltSel, wrapSel, nodeSel){\n"
        "  function pickFromMain(){\n"
        "    try{\n"
        "      var mainText = document.querySelector(mainSel);\n"
        "      var symbolEl = document.querySelector(symSel);\n"
        "      if (mainText && mainText.textContent){\n"
        "        var sym = symbolEl && symbolEl.textContent ? symbolEl.textContent.trim() : '¥';\n"
        "        var txt = mainText.textContent.trim();\n"
        "        if (txt) return sym + txt;\n"
        "      }\n"
        "    }catch(e){}\n"
        "    return '';\n"
        "  }\n"
        "  function pickFromAlts(){\n"
        "    try{\n"
        "      alts = Array.isArray(alts) ? alts : [];\n"
        "      for (var i=0; i<alts.length; i++){\n"
        "        var el = document.querySelector(alts[i]);\n"
        "        if (el && el.textContent){\n"
        "          var t = el.textContent.trim();\n"
        "          if (t){\n"
        "            if (t.indexOf('¥') !== -1 || t.indexOf('￥') !== -1) return t;\n"
        "            return '¥' + t;\n"
        "          }\n"
        "        }\n"
        "      }\n"
        "    }catch(e){}\n"
        "    return '';\n"
        "  }\n"
        "  function pickFromBelt(){\n"
        "    try{\n"
        "      var belt = document.querySelector(beltSel);\n"
        "      if (!belt) return '';\n"
        "      var container = belt.querySelector(wrapSel) || belt;\n"
        "      var nodes = container.querySelectorAll(nodeSel);\n"
        "      for (var j=0; j<nodes.length; j++){\n"
        "        var tt = (nodes[j].textContent || '').trim();\n"
        "        if (tt && /\\d/.test(tt)){\n"
        "          if (tt.indexOf('¥') !== -1 || tt.indexOf('￥') !== -1) return tt;\n"
        "          return '¥' + tt;\n"
        "        }\n"
        "      }\n"
        "    }catch(e){}\n"
        "    return '';\n"
        "  }\n"
        "  return pickFromMain() || pickFromAlts() || pickFromBelt();\n"
        "})(arguments[0], arguments[1], arguments[2], arguments[3], arguments[4], arguments[5]);"
    )

    t_price0 = time.perf_counter()
    end_time = time.perf_counter() + 0.3
    last = ""
    while time.perf_counter() < end_time:
        try:
            price = (
                driver.execute_script(
                    js,
                    PRICE_MAIN_TEXT,
                    PRICE_SYMBOL,
                    PRICE_ALT_SELECTORS,
                    "[class*='beltPrice']",
                    "[class*='priceWrap']",
                    "[class*='text'], [class*='number']"
                )
                or ""
            ).strip()
            if price and any(ch.isdigit() for ch in price):
                return normalize_price_text(price)
            last = price
        except Exception:
            pass
        time.sleep(0.06)
    price_final = normalize_price_text(last) or "未获取到价格"
    log_debug(f"取价耗时 {(time.perf_counter() - t_price0)*1000:.0f}ms，结果 {price_final}")
    return price_final


def get_main_image_url(driver) -> str:
    """获取当前主图的大图 URL（方案A）。
    优先返回主图 <img> 的 src（保持原样，包含 .webp 等后缀）；若未找到，则回退到放大镜容器的背景图 URL。
    """
    js = (
        "return (function(imgSel, zoomSel){\n"
        "  function extractFromBg(bg){\n"
        "    try{ if(!bg) return ''; }catch(e){ return ''; }\n"
        "    var m = String(bg).match(/url\\([\"']?(.*?)[\"']?\\)/i);\n"
        "    if (m && m[1]) return (m[1] || '').trim();\n"
        "    return '';\n"
        "  }\n"
        "  // 1) 先取主图 <img> 的 src\n"
        "  try{\n"
        "    var img = document.querySelector(imgSel);\n"
        "    if (img && img.src){ return (img.src || '').trim(); }\n"
        "  }catch(e){}\n"
        "  // 2) 兜底：放大镜容器背景图\n"
        "  try{\n"
        "    var zoom = document.querySelector(zoomSel);\n"
        "    if (zoom){\n"
        "      var cs = window.getComputedStyle ? window.getComputedStyle(zoom) : null;\n"
        "      var bg = (zoom.style && zoom.style.backgroundImage) || (cs && cs.backgroundImage) || '';\n"
        "      var hi = extractFromBg(bg);\n"
        "      if (hi) return hi;\n"
        "    }\n"
        "  }catch(e){}\n"
        "  return '';\n"
        "})(arguments[0], arguments[1]);"
    )

    try:
        url = (
            driver.execute_script(js, MAIN_PIC_IMG_SELECTOR, ZOOM_IMG_DIV_SELECTOR) or ""
        ).strip()
        return url
    except Exception:
        return ""


def compute_total_combinations(sku_dimensions: List[SkuDimension]) -> int:
    total = 1
    for dim in sku_dimensions:
        total *= len(dim.options)
    return total


def print_dimensions_summary(sku_dimensions: List[SkuDimension]) -> None:
    total_combinations = compute_total_combinations(sku_dimensions)
    print(f"[信息] 检测到 {len(sku_dimensions)} 个维度，总共 {total_combinations} 个组合")
    for i, dim in enumerate(sku_dimensions):
        print(f"  维度{i+1}: {dim.name} ({len(dim.options)}个选项)")


def write_dimensions_structure_log(sku_dimensions: List[SkuDimension]) -> None:
    """将所有维度及其选项结构输出到 log/sku维度选项结构.log（每次运行覆盖写入）。"""
    log_dir = Path(__file__).resolve().parent / "log"
    log_dir.mkdir(parents=True, exist_ok=True)
    path = log_dir / "sku维度选项结构.log"
    lines: List[str] = []
    lines.append("# SKU维度与选项结构\n")
    for i, dim in enumerate(sku_dimensions, 1):
        lines.append(f"维度{i}: {dim.name}\n")
        for j, opt in enumerate(dim.options, 1):
            lines.append(f"  - 选项{j}: {opt.text} (vid={opt.vid})\n")
        lines.append("\n")
    with open(path, "w", encoding="utf-8") as f:
        f.writelines(lines)

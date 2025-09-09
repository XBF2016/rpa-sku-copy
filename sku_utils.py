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
    """获取当前主图的大图 URL（增强版）。
    优先顺序：
    1) 主图 <img> 的 currentSrc（适配 srcset）
    2) 主图 <img> 的 src（直接地址，可能为 .webp 等）
    3) 主图 <img> 的 srcset/placeholder/data-src/data-ks-lazyload、<picture> 下 <source> 的 srcset
    4) 兜底：放大镜容器（.js-image-zoom__zoomed-image）的 background-image URL
    仅返回以 http/https 开头的链接，避免 data: 等占位符。
    """
    # 先尝试将主图区域滚动到视口中，增加懒加载触发概率
    try:
        driver.execute_script(
            "try{var w=document.querySelector(\"[class*='mainPicWrap']\"); if(w){w.scrollIntoView({block:'center',inline:'center'});} }catch(e){}"
        )
    except Exception:
        pass

    js = (
        "return (function(imgSel, zoomSel){\n"
        "  function isHttp(u){ try{ return typeof u === 'string' && /^https?:\\/\\//i.test(u); }catch(e){ return false; } }\n"
        "  function tryZoomPreload(){\n"
        "    try{\n"
        "      var wrap = document.querySelector(\"[class*='mainPicWrap']\");\n"
        "      if(!wrap) return;\n"
        "      var rect = wrap.getBoundingClientRect();\n"
        "      var cx = rect.left + rect.width * 0.6;\n"
        "      var cy = rect.top + rect.height * 0.6;\n"
        "      ['mouseenter','mouseover','mousemove'].forEach(function(tp){\n"
        "        try{ wrap.dispatchEvent(new MouseEvent(tp, {bubbles:true, clientX: cx, clientY: cy, view: window})); }catch(e){}\n"
        "      });\n"
        "    }catch(e){}\n"
        "  }\n"
        "  function fromImg(img){\n"
        "    if(!img) return '';\n"
        "    try{ var u = img.currentSrc || ''; if(isHttp(u)) return u.trim(); }catch(e){}\n"
        "    try{ var u2 = img.getAttribute('src') || ''; if(isHttp(u2)) return u2.trim(); }catch(e){}\n"
        "    try{ var u3 = img.src || ''; if(isHttp(u3)) return u3.trim(); }catch(e){}\n"
        "    try{\n"
        "      var ss = img.getAttribute('srcset') || '';\n"
        "      if(ss){ var first = ss.split(',')[0].trim().split(' ')[0].trim(); if(isHttp(first)) return first; }\n"
        "    }catch(e){}\n"
        "    try{\n"
        "      var lazy = img.getAttribute('data-src') || img.getAttribute('data-original') || img.getAttribute('data-lazyload') || img.getAttribute('data-lazy') || img.getAttribute('data-srcset') || img.getAttribute('data-ks-lazyload') || img.getAttribute('placeholder') || '';\n"
        "      if(isHttp(lazy)) return lazy.trim();\n"
        "    }catch(e){}\n"
        "    try{\n"
        "      var pic = img.closest('picture');\n"
        "      if (pic){\n"
        "        var s = pic.querySelector('source[srcset]');\n"
        "        if (s){ var ss2 = s.getAttribute('srcset') || ''; var first2 = ss2.split(',')[0].trim().split(' ')[0].trim(); if(isHttp(first2)) return first2; }\n"
        "      }\n"
        "    }catch(e){}\n"
        "    return '';\n"
        "  }\n"
        "  function extractFromBg(bg){\n"
        "    try{ if(!bg) return ''; }catch(e){ return ''; }\n"
        "    var m = String(bg).match(/url\\([\"']?(.*?)[\"']?\\)/i);\n"
        "    if (m && m[1]) return (m[1] || '').trim();\n"
        "    return '';\n"
        "  }\n"
        "  // 1) 优先从主图 <img> 提取（兼容 srcset/懒加载占位符）\n"
        "  try{\n"
        "    var img = document.querySelector(imgSel);\n"
        "    if (!img){ var wrap = document.querySelector(\"[class*='mainPicWrap']\"); if (wrap) img = wrap.querySelector('img'); }\n"
        "    var u = fromImg(img);\n"
        "    if (u) return u;\n"
        "  }catch(e){}\n"
        "  // 2) 兜底：取放大镜容器的背景图 URL（必要时尝试预加载）\n"
        "  try{\n"
        "    tryZoomPreload();\n"
        "    var zoom = document.querySelector(zoomSel);\n"
        "    if (zoom){\n"
        "      var cs = window.getComputedStyle ? window.getComputedStyle(zoom) : null;\n"
        "      var bg = (zoom.style && zoom.style.backgroundImage) || (cs && cs.backgroundImage) || '';\n"
        "      var hi = extractFromBg(bg);\n"
        "      if (isHttp(hi)) return hi;\n"
        "    }\n"
        "  }catch(e){}\n"
        "  // 3) 进一步兜底：若主图区域自身使用 background-image\n"
        "  try{\n"
        "    var wrap2 = document.querySelector(\"[class*='mainPicWrap']\");\n"
        "    if (wrap2){\n"
        "      var cs2 = window.getComputedStyle ? window.getComputedStyle(wrap2) : null;\n"
        "      var bg2 = (wrap2.style && wrap2.style.backgroundImage) || (cs2 && cs2.backgroundImage) || '';\n"
        "      var hi2 = extractFromBg(bg2);\n"
        "      if (isHttp(hi2)) return hi2;\n"
        "    }\n"
        "  }catch(e){}\n"
        "  // 4) 最终兜底：在全局范围尝试若干候选选择器\n"
        "  try{\n"
        "    var candSelectors = [\n"
        "      imgSel,\n"
        "      \"[class*='mainPicWrap'] img\",\n"
        "      \"img[src*='alicdn.com']\",\n"
        "      \"img[src*='gw.alicdn.com']\",\n"
        "      \"img[src*='/bao/uploaded/']\",\n"
        "      \"img[src*='/imgextra/']\"\n"
        "    ];\n"
        "    var tried = new Set();\n"
        "    for (var k=0; k<candSelectors.length; k++){\n"
        "      var sel = candSelectors[k];\n"
        "      if (tried.has(sel)) continue; tried.add(sel);\n"
        "      var nodes = document.querySelectorAll(sel);\n"
        "      for (var i=0; i<nodes.length; i++){\n"
        "        var u = fromImg(nodes[i]);\n"
        "        if (isHttp(u)) return u;\n"
        "      }\n"
        "    }\n"
        "  }catch(e){}\n"
        "  // 5) meta/link 兜底：如 og:image / image_src\n"
        "  try{\n"
        "    var m = document.querySelector(\"meta[property='og:image']\") || document.querySelector(\"meta[name='og:image']\") || document.querySelector(\"meta[property='og:image:secure_url']\");\n"
        "    if (m){ var c = m.getAttribute('content') || ''; if (isHttp(c)) return c; }\n"
        "    var l = document.querySelector(\"link[rel='image_src']\");\n"
        "    if (l){ var h = l.getAttribute('href') || ''; if (isHttp(h)) return h; }\n"
        "  }catch(e){}\n"
        "  return '';\n"
        "})(arguments[0], arguments[1]);"
    )
    # Python 端短轮询：等待主图URL在点击SKU后稳定可用（最多 ~0.8s）
    end_time = time.perf_counter() + 0.8
    last = ""
    while time.perf_counter() < end_time:
        try:
            url = (
                driver.execute_script(js, MAIN_PIC_IMG_SELECTOR, ZOOM_IMG_DIV_SELECTOR) or ""
            ).strip()
            # 兼容以 // 开头的协议相对地址
            if url.startswith("//"):
                url = "https:" + url
            if url.startswith("http://") or url.startswith("https://"):
                return url
            last = url
        except Exception as e:
            last = ""
        time.sleep(0.05)
    # 若超时仍无 http 链接，输出详细调试信息
    try:
        if last:
            log_debug(f"主图URL候选(非http): {last}")
        dbg = driver.execute_script(
            "return (function(imgSel, zoomSel){\n"
            "  try{\n"
            "    var wrap=document.querySelector(\"[class*='mainPicWrap']\");\n"
            "    var img = wrap? wrap.querySelector('img') : document.querySelector(imgSel);\n"
            "    var zoom=document.querySelector(zoomSel);\n"
            "    var o=[];\n"
            "    o.push('wrap存在='+(!!wrap));\n"
            "    o.push('img存在='+(!!img));\n"
            "    if(img){\n"
            "      o.push('img.src='+(img.getAttribute('src')||''));\n"
            "      o.push('img.currentSrc='+(img.currentSrc||''));\n"
            "      o.push('img.srcset='+(img.getAttribute('srcset')||''));\n"
            "      o.push('img.data-src='+(img.getAttribute('data-src')||''));\n"
            "      o.push('img.placeholder='+(img.getAttribute('placeholder')||''));\n"
            "    }\n"
            "    o.push('zoom存在='+(!!zoom));\n"
            "    if(zoom){\n"
            "      var cs=window.getComputedStyle?window.getComputedStyle(zoom):null;\n"
            "      var bg=(zoom.style&&zoom.style.backgroundImage)||(cs&&cs.backgroundImage)||'';\n"
            "      o.push('zoom.bg='+bg);\n"
            "    }\n"
            "    return o.join(' | ');\n"
            "  }catch(e){ return '调试收集异常'; }\n"
            "})(arguments[0], arguments[1]);",
            MAIN_PIC_IMG_SELECTOR,
            ZOOM_IMG_DIV_SELECTOR,
        ) or ''
        if dbg:
            log_debug(f"主图调试: {dbg}")
    except Exception:
        pass
    return last if (last.startswith("http://") or last.startswith("https://")) else ""


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

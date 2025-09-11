import os
from pathlib import Path

# 统一的公共工具：配置路径、日志、调试、文本规范化、读取商品链接


def project_root() -> Path:
    """项目根目录（以当前文件所在目录为根）。"""
    return Path(__file__).resolve().parent


def conf_path(name: str) -> Path:
    """返回 conf 目录下的配置文件路径。"""
    return project_root() / "conf" / name


# 控制台/日志文本规范化（解决 Windows 控制台 GBK 下 '¥' 无法编码的问题）
def normalize_price_text(txt: str) -> str:
    try:
        if not isinstance(txt, str):
            txt = str(txt)
    except Exception:
        txt = ""
    # 将 '¥'(U+00A5) 统一替换为 '￥'(U+FFE5)
    return (txt or "").replace("¥", "￥").strip()


def debug_on() -> bool:
    """调试开关：通过环境变量 DEBUG_RPA（1/true/yes/y/on）。"""
    try:
        v = os.environ.get("DEBUG_RPA", "").strip().lower()
        return v in ("1", "true", "yes", "y", "on")
    except Exception:
        return False


def append_to_log(message: str) -> None:
    """向 log/sku维度及选项.log 追加一行带时间戳的日志。"""
    import datetime
    log_dir = project_root() / "log"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / "sku维度及选项.log"

    timestamp = datetime.datetime.now().strftime("%H:%M:%S")
    log_line = f"[{timestamp}] {message}\n"

    try:
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(log_line)
            f.flush()
    except UnicodeEncodeError:
        clean_message = message.encode("utf-8", errors="ignore").decode("utf-8")
        clean_log_line = f"[{timestamp}] {clean_message}\n"
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(clean_log_line)
            f.flush()


def log_debug(message: str) -> None:
    """输出调试信息（控制台 + 日志文件），仅在开启调试时生效。"""
    if debug_on():
        print(f"[调试] {message}")
        try:
            append_to_log(f"调试: {message}")
        except Exception:
            pass


def read_browser_path() -> str:
    """从 conf/browser.txt 读取 Edge 可执行文件路径；若为空则回退为 'msedge.exe'。"""
    path_file = conf_path("browser.txt")
    try:
        content = path_file.read_text(encoding="utf-8").strip().strip('"')
        if content:
            return content
    except Exception:
        pass
    return "msedge.exe"


def read_product_url() -> str:
    """从 conf/product-url.txt 读取商品链接（取第一条有效行）。
    规则：忽略空行、忽略以#开头的注释行，取第一条以 http/https 开头的链接。
    """
    path_file = conf_path("product-url.txt")
    if not path_file.exists():
        raise ValueError("未找到 conf/product-url.txt，请在 conf 目录提供该文件")
    for line in path_file.read_text(encoding="utf-8").splitlines():
        url = line.strip()
        if not url:
            continue
        if url.startswith('#'):
            continue
        if url.startswith("http://") or url.startswith("https://"):
            return url
        raise ValueError("product-url.txt 中存在非注释的第一条内容不是有效的 http/https 链接")
    raise ValueError("product-url.txt 中没有可用的链接内容（已忽略空行与注释行）")

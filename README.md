# rpa-sku-copy

本项目用于自动遍历商品页面的 SKU 组合，抓取每个组合的价格并导出结果。

## 新的项目结构

- `main.py`：RPA 任务入口（包含 `@task traverse_all_sku_combinations` 与 `main()`）。
- `common.py`：通用工具（配置路径、日志、调试、价格文本规范化、读取商品链接）。
- `browser_utils.py`：浏览器/Driver 管理（清理进程、查找本地 `msedgedriver.exe`、初始化 WebDriver、打开页面）。
- `sku_utils.py`：页面选择器常量、SKU 数据模型、SKU 解析、取价、维度概要与结构日志输出。
- `traversal.py`：遍历相关逻辑（生成组合、首选当前选择、限制数量、点击选择、遍历收集）。
- `io_utils.py`：导出 Excel。
- `conf/`：所有配置文件目录（`conda.yaml`、`robot.yaml`、`product-url.txt`、`browser.txt`）。
- `driver/`：浏览器驱动目录（`msedgedriver.exe`）。
- `output/`：导出结果目录（Excel）。
- `log/`：运行日志目录。

说明：原 `tasks.py` 的逻辑已拆分到上述模块中，推荐使用 `main.py` 作为新的入口。

## 配置文件说明（conf/）
- `product-url.txt`：第一行有效的商品 URL（忽略空行与 `#` 注释行）。
- `browser.txt`：Edge 浏览器可执行文件路径（例如：`C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe`），若留空则回退 `msedge.exe`。
- `robot.yaml`：Robocorp 任务描述，已指向上级目录的 `main.py`：
  - `shell: python -m robocorp.tasks run ../main.py -t traverse_all_sku_combinations`
  - `condaConfigFile: conda.yaml`
  - `artifactsDir: ../output`（可按需修改）。运行时 RCC 会注入环境变量 `ROBOT_ARTIFACTS` 指向该目录，代码将优先使用此目录作为导出根目录。
  - `PATH: ['..']`，`PYTHONPATH: ['..']`
- `conda.yaml`：Python 环境依赖（包含 robocorp-tasks、selenium、openpyxl 等）。
  - 现已增加 `pillow` 依赖，用于在 Excel 中嵌入图片。
  - 依赖精简：移除 `webdriver-manager`（不再联网下载驱动），通过本地路径自动查找 `msedgedriver.exe`（见 `browser_utils.find_msedgedriver_path()`）。

## 浏览器驱动（driver/）
- 将 `msedgedriver.exe` 放置到 `driver/` 目录（默认优先查找此处）。
- 也支持通过环境变量 `MSEDGEDRIVER` 指向目录或文件，或与 `msedge.exe` 同目录、或常见安装路径、或 `PATH` 解析。

## 运行方式
1. 通过 Robocorp Tasks 运行（推荐本地直跑）：
   ```powershell
   python -m robocorp.tasks run main.py -t traverse_all_sku_combinations
   ```
2. 通过 `robot.yaml` 运行（需要 `rcc` 环境）：
   ```powershell
   # 从项目根目录执行
   rcc run -r conf\robot.yaml
   ```

### 安装 RCC（Windows）
如果本机还未安装 `rcc`，可按以下步骤安装（无需管理员权限）：

```powershell
# 1) 创建用户级 bin 目录（若已存在会跳过）
New-Item -ItemType Directory -Force "$env:USERPROFILE\bin" | Out-Null

# 2) 下载 rcc.exe 到该目录（使用官方下载地址）
Invoke-WebRequest -UseBasicParsing -Uri "https://downloads.robocorp.com/rcc/releases/latest/windows64/rcc.exe" -OutFile "$env:USERPROFILE\bin\rcc.exe"

# 3) 将该目录加入当前会话的 PATH（仅对本次 PowerShell 会话生效）
$env:PATH = "$env:USERPROFILE\bin;" + $env:PATH

# 4) 验证安装是否成功
rcc --version
```

安装完成后，即可通过：

```powershell
$env:MAX_COMBOS=5; rcc run -r conf\robot.yaml
```

来仅跑 5 条进行快速验证。

## 其他
- 控制台/日志统一中文输出；价格文本统一替换 `¥` 为 `￥` 以避免 GBK 编码问题。
- 可通过环境变量 `DEBUG_RPA=1` 打开调试日志；`MAX_COMBOS=N` 可限制前 N 个组合用于快速验证。
- 导出的 Excel：
  - 新增“图片”与“图片链接”两列（均位于“价格”之前）。
    - “图片”列：仅插入图片，不写入链接文本。图片会被下载并按【图片列列宽】等比例缩放后，锚定到该单元格内，尽量不溢出到相邻单元格；行高将根据缩放后的图片高度自动设置，达到“看起来就在单元格里”的效果。默认图片预览宽度约 160px（可在 `io_utils.py` 中调整）。
    - “图片链接”列：始终写入可点击的图片 URL（便于复制粘贴或外部打开大图），默认列宽较宽（90）。
  - 全表样式：所有单元格均设置为“水平居中 + 垂直居中 + 自动换行”。
  - 列宽：除“图片/图片链接”外的文本列，会根据内容长度做近似自适应列宽（限定在 10~40 之间，避免过宽/过窄）。
  - “价格”列表头识别为“价格”时会以纯数值写入（自动去掉货币符号与千分位），方便后续做数值计算与筛选。
  - 图片格式兼容性：在嵌入到 Excel 前，会将下载的图片统一转为 PNG（通过 Pillow），以避免 WEBP 等格式在 Excel 中无法直接嵌入导致的错误；若转换失败，会保留图片链接文本。
  - 主图提取：`sku_utils.get_main_image_url()` 已增强：
    - 顺序：`<img>.currentSrc/src/srcset/placeholder/data-src/data-ks-lazyload/...`、`<picture><source srcset>`；
    - 兜底：放大镜容器 `.js-image-zoom__zoomed-image` 的 `background-image`；`[class*='mainPicWrap']` 自身的 `background-image`；
    - 最终兜底：在全局尝试若干主图候选选择器，必要时使用 `meta[property='og:image']` / `link[rel=image_src]`；
    - 兼容以 `//` 开头的协议相对地址（自动补全为 `https:`）；
    - 在点击SKU后做短轮询等待（~0.8s）并尝试触发一次主图区域的悬停，以促使放大镜背景图预加载。

### 导出文件存放路径
- 若通过 `robot.yaml`（RCC）运行：结果文件将保存为：`<artifactsDir>/[店铺名]商品名/result.xlsx`，其中 `<artifactsDir>` 由 `conf/robot.yaml` 的 `artifactsDir` 控制（默认 `../output`）。
- 若直接本地运行（未通过 RCC/robot.yaml）：结果文件将保存为：`output/[店铺名]商品名/result.xlsx`
- “店铺名/商品名”的解析基于元素示例：
  - 商品名：`元素示例/商品名.html`（选择器：`[class*='mainTitle--']`）
  - 店铺名：`元素示例/店铺名.html`（选择器：`[class*='shopName--']`）
- 文件夹名称会自动清理非法字符（例如 `\/:*?"<>|` 等），并做适当长度截断，确保在 Windows 下可用。

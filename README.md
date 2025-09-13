## 登录态复用（避免“登录过期”）
- 程序在启动 Edge 时会尽量复用本机已登录的浏览器用户数据目录，避免出现“登录过期”。
- 默认策略：
  - 若未设置环境变量，会自动尝试使用 `%LOCALAPPDATA%\\Microsoft\\Edge\\User Data` 作为 `--user-data-dir`，并使用 `--profile-directory=Default`。
  - 若你平时登录抖店使用的不是 `Default` 配置（例如 `Profile 1`/`Profile 2`），请在运行前显式设置：

  ```powershell
  $env:EDGE_USER_DATA_DIR="$env:LOCALAPPDATA\Microsoft\Edge\User Data"
  $env:EDGE_PROFILE="Profile 1"   # 或者 Default / Profile 2 等
  $env:MAX_COMBOS=5; python -m robocorp.tasks run .\main.py -t create_douyin_spec_dimensions
  ```

- 若仍然看见“登录过期”，请在浏览器页面内手动完成登录；程序会循环等待，直到检测到“添加规格类型”按钮或超时（默认 10 分钟）。
# 抖店 RPA：创建规格维度 + 录入选项 + 填写价格 + 保存草稿

本 RPA 用于在抖店“商品上架草稿页”中，自动创建商品规格的“维度”，并在每个维度下录入所有“选项”，随后根据 `output/<子目录>/result.yml` 的 `combos` 中的价格顺序，为 SKU 表格的“价格”列依次填写价格（幂等补充），最后点击“保存草稿”。

- 页面元素严格参考 `元素示例/` 目录：
  - `元素示例/添加规格类型.html`
  - `元素示例/规格类型下拉按钮.html`
  - `元素示例/规格类型下拉列表.html`
- 会创建维度并录入每个维度的所有选项，然后填写 SKU 价格，最后保存草稿。
- 全部中文注释、中文日志、中文报错。

## 目录结构（相关文件）
- `main.py`：任务入口，仅负责任务编排与入口（调用下述模块）
- `config.py`：路径常量与配置读取（浏览器路径、商品链接、规格最简解析：支持从 output/<子目录>/result.yml 的 specs 读取）
- `driver_utils.py`：浏览器启动与稳健导航、登录等待等通用流程
- `douyin_actions.py`：抖店页面元素与动作（创建维度、录入选项、保存草稿等）
- `conf/browser.txt`：Edge 浏览器可执行文件路径（例如：`C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe`）
- `conf/product-url.txt`：商品上架草稿页链接（若 `conf/douyin/product-url.txt` 不存在，则回退使用本文件）
- `conf/douyin/product-url.txt`：商品上架草稿页链接（优先读取）
- `conf/规格.yml`：规格配置文件，仅使用其“顶层键名”作为需要创建的维度列表
- `driver/msedgedriver.exe`：EdgeDriver（若不存在，将尝试使用系统路径中的 WebDriver）
- `元素示例/`：抖店页面元素示例（用于选择器定位的依据）

## 规格与价格配置来源（output/<子目录>/result.yml）
不引入额外依赖，采用最简解析：
从 `output/` 下“最新修改时间”的子目录中的 `result.yml` 读取 `specs` 段，
其中 `specs` 的顶层键名作为“维度”，该键名下缩进的 `- 项` 作为该维度的“选项”。例如：

```yaml
颜色分类:
  - 红色
  - 蓝色
适用床垫尺寸:
  - 1500mm*2000mm
  - 1800mm*2000mm
家具结构:
  - 框架结构
  - 榫卯结构
```

上述示例中，将会创建维度：`颜色分类`、`适用床垫尺寸`、`家具结构`，并在各自维度下依次输入所有列出的选项。
此外，程序会解析：

```yaml
dims:
  - 颜色分类
  - 适用床垫尺寸
  - 家具结构
combos:
  - [0, 0, 0, 1431.0, 0]
  - [0, 0, 1, 2241.0, 0]
```

其中 `combos` 的每一行，前 `len(dims)` 个为选项索引，紧随其后的位置为价格（本项目按示例取第 `len(dims)` 位作为价格），其后为图片索引（可忽略）。程序会按 `combos` 的顺序为 SKU 表格的“价格”输入框依次填写价格；若 `combos` 为空或价格为 `null`/小于 0.01，将跳过该项。
若 `output/` 下没有任何 `result.yml`，程序会报错提示请先生成规格数据。

## 运行方式
- 方式一（直接用 Python 运行）：

  在 Windows PowerShell 中（按照项目规范，带上演示用的环境变量）：

  ```powershell
  $env:MAX_COMBOS=5; python -m robocorp.tasks run .\main.py -t create_douyin_spec_dimensions
  ```

   说明：
  - 若 `result.yml` 中存在 `combos`，程序会尝试填充价格；否则会跳过价格填写。
  - 若页面需要登录，请在浏览器打开后先完成登录；程序会等待“添加规格类型”按钮出现后继续。

- 方式二（通过 RCC/robot.yaml）

  `conf/robot.yaml` 当前指向 `../main.py`。若直接使用 rcc 运行该任务，可执行：

  ```powershell
  rcc run
  ```

  （如需单独筛选任务，可在命令中追加 `-t create_douyin_spec_dimensions`。）

## 元素选择器说明（严格按“元素示例/”）
- “添加规格类型”按钮：通过包含 `ecom-g-btn-dashed` 的按钮且文案包含“添加规格类型”定位。
- “规格类型下拉按钮”：通过占位文案“请选择规格类型”的下拉控件定位。
- “规格类型下拉列表”：通过 `ecom-g-select-dropdown` 且非 `ecom-g-select-dropdown-hidden` 定位当前可见的下拉。
- “下拉列表项”：通过项内容容器 `ecom-g-select-item-option-content` 与精确文本匹配定位。
- “创建类型”链接：通过下拉底部文案“创建类型”的链接定位。
- “请输入规格类型”输入框：通过 placeholder 包含“请输入规格类型”定位（点击“创建类型”后出现）。
- “已添加维度”识别：通过页面中的 `ecom-g-select-selection-item` 文本/标题提取。
- “规格值输入框”：参考 `元素示例/商品规格区域.html`，容器 `id="skuValue-<维度名>"` 下的 input，且 placeholder 为“请输入<维度名>”；支持全局退化按 placeholder 查找。
- “保存草稿”按钮：参考 `元素示例/保存草稿按钮.html`，通过按钮文案“保存草稿”定位。
 - “SKU表格与价格输入框”：参考 `元素示例/sku表格区域.html` 与 `元素示例/sku价格输入框.html`，在虚拟表格 `ecom-g-table-tbody-virtual ecom-g-table-tbody` 容器内，按行查找 `attr-column-field_price` 列下的 `input.ecom-g-input-number-input` 输入框，依次输入价格。

## 运行日志与幂等
- 程序会在开始时打印将要创建的维度列表；每一步操作均有中文提示。
- 对已存在的维度会自动跳过，避免重复创建（幂等）。
 - 价格填写遵循幂等：若对应输入框已有大于 0 的值，则跳过不覆盖；对 `null` 或小于 0.01 的价格也跳过。

- 若仍然看见“登录过期”，请在浏览器页面内手动完成登录；程序会循环等待，直到检测到“添加规格类型”按钮或超时（默认 10 分钟）。

## 稳定性处理（避免启动/导航异常）
- 启动前自动关闭 Edge：程序会在启动前调用 `browser_utils.kill_edge_processes()` 强制结束 `msedge.exe/msedgewebview2.exe`，避免“用户数据目录被占用”导致的崩溃。
- 导航备用方案：若 `driver.get(url)` 后仍停留在空白页，将尝试 `window.open(url)` 新标签并切换，仍不行则用 `location.href = url` 脚本导航。
- 驱动与附加模式回退：若本地 EdgeDriver 启动失败（DevToolsActivePort/failed to start/crashed），程序会自动回退为 Selenium Manager 自动匹配；若仍失败，将自动启动带 `--remote-debugging-port` 的 Edge 并“附加”连接。
- 用户数据目录占用自动回退：若报错提示用户数据目录被占用（already in use），程序会自动改用临时用户数据目录（`temp/edge-user-data-<时间戳>`）重新启动，避免冲突；仍失败时再尝试“附加模式”。

## 性能与速度
- 规格选项录入采用“极速批量输入”模式：先在“新增用输入框”快速连输（每项仅约 20~30ms 极短等待），一轮结束后统一校验缺失项并补录一次，显著减少每项的等待时间。
- 若你的页面网络或设备较慢，导致个别选项首次未识别，也会在补录阶段自动补齐；如仍有漏项，可再次运行任务，程序会幂等跳过已存在内容，仅补缺。

## 注意事项
- 首次运行可能需要登录抖店后台：程序会在未检测到“添加规格类型”按钮时提示等待登录，然后继续。
- 若 `conf/douyin/product-url.txt` 不存在，将自动回退使用 `conf/product-url.txt`。
- `driver/msedgedriver.exe` 建议放置在 `driver/` 目录下以避免系统路径差异。

## 变更说明
- 完成重构与拆分：
  - `main.py`：精简为任务编排入口
  - `config.py`：新增，承载路径常量与配置读取
  - `driver_utils.py`：新增，承载浏览器启动、导航与登录等待
  - `douyin_actions.py`：新增，承载抖店页面元素与操作
- `conf/robot.yaml`：任务入口保持不变（仍指向 `../main.py`）。
- `README.md`：同步更新目录结构与说明。

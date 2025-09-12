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
# 抖店 RPA：创建商品规格维度（最小实现）

本 RPA 用于在抖店“商品上架草稿页”中，自动创建商品规格的“维度”（不创建维度下的选项）。

- 页面元素严格参考 `元素示例/` 目录：
  - `元素示例/添加规格类型.html`
  - `元素示例/规格类型下拉按钮.html`
  - `元素示例/规格类型下拉列表.html`
- 仅添加维度，暂不添加维度选项，满足本次最小需求。
- 全部中文注释、中文日志、中文报错。

## 目录结构（相关文件）
- `main.py`：任务入口，定义任务 `create_douyin_spec_dimensions`
- `conf/browser.txt`：Edge 浏览器可执行文件路径（例如：`C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe`）
- `conf/product-url.txt`：商品上架草稿页链接（若 `conf/douyin/product-url.txt` 不存在，则回退使用本文件）
- `conf/douyin/product-url.txt`：商品上架草稿页链接（优先读取）
- `conf/规格.yml`：规格配置文件，仅使用其“顶层键名”作为需要创建的维度列表
- `driver/msedgedriver.exe`：EdgeDriver（若不存在，将尝试使用系统路径中的 WebDriver）
- `元素示例/`：抖店页面元素示例（用于选择器定位的依据）

## 规格配置（conf/规格.yml）
为减少依赖与复杂度，本实现仅解析 YAML 的“顶层键名”作为维度名称。例如：

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

上述示例中，最终会创建的“维度”为：`颜色分类`、`适用床垫尺寸`、`家具结构`（不会创建选项）。

## 运行方式
- 方式一（直接用 Python 运行）：

  在 Windows PowerShell 中（按照项目规范，带上演示用的环境变量）：

  ```powershell
  $env:MAX_COMBOS=5; python -m robocorp.tasks run .\main.py -t create_douyin_spec_dimensions
  ```

  说明：
  - 该任务不涉及 SKU 组合遍历，因此 `MAX_COMBOS` 仅作为统一运行风格的占位变量。
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

## 运行日志与幂等
- 程序会在开始时打印将要创建的维度列表；每一步操作均有中文提示。
- 对已存在的维度会自动跳过，避免重复创建（幂等）。

## 注意事项
- 首次运行可能需要登录抖店后台：程序会在未检测到“添加规格类型”按钮时提示等待登录，然后继续。
- 若 `conf/douyin/product-url.txt` 不存在，将自动回退使用 `conf/product-url.txt`。
- `driver/msedgedriver.exe` 建议放置在 `driver/` 目录下以避免系统路径差异。

## 变更说明
- 新增 `main.py`（仅一个文件），实现“抖店-创建规格维度”的最小功能；
- 未修改 `conf/robot.yaml` 的现有任务定义，保持原有结构不变；
- 本 README 为本功能的使用说明，确保文档与实现同步，不出现过时描述。

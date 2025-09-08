@echo off
chcp 65001 >nul
echo ========================================
echo           RPA SKU 复制工具
echo ========================================
echo.
echo 此脚本将依次执行以下任务：
echo 1. 打开Edge浏览器（保持登录状态）
echo 2. 打开商品页面
echo 3. 提取SKU维度和选项
echo 4. 遍历所有SKU组合获取价格
echo.
echo 注意：请确保已经：
echo - 安装了所需的Python环境和依赖
echo - 在product-url.txt中配置了商品链接
echo - Edge浏览器已登录淘宝账号
echo.
set /p confirm=确认要开始运行RPA吗？(Y/N): 

if /i "%confirm%" NEQ "Y" (
    echo 操作已取消
    pause
    exit /b 1
)

echo.
echo ========================================
echo 开始执行RPA任务...
echo ========================================

echo [步骤 1/4] 打开Edge浏览器（保持登录状态）...
python -m robocorp.tasks run tasks.py -t open_edge_logged_in
if %errorlevel% neq 0 (
    echo [错误] 步骤1执行失败
    pause
    exit /b 1
)

echo.
echo [步骤 2/4] 打开商品页面...
python -m robocorp.tasks run tasks.py -t open_product_page
if %errorlevel% neq 0 (
    echo [错误] 步骤2执行失败
    pause
    exit /b 1
)

echo.
echo [步骤 3/4] 提取SKU维度和选项...
python -m robocorp.tasks run tasks.py -t extract_sku_dimensions
if %errorlevel% neq 0 (
    echo [错误] 步骤3执行失败
    pause
    exit /b 1
)

echo.
echo [步骤 4/4] 遍历所有SKU组合获取价格...
python -m robocorp.tasks run tasks.py -t traverse_all_sku_combinations
if %errorlevel% neq 0 (
    echo [错误] 步骤4执行失败
    pause
    exit /b 1
)

echo.
echo ========================================
echo 所有任务执行完成！
echo ========================================
echo.
echo 结果文件位置：
echo - SKU维度和选项: log\sku维度及选项.log  
echo - SKU组合价格: log\sku维度选项组合.log
echo.
echo 按任意键退出...
pause >nul

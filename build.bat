@echo off
chcp 65001 >nul
echo ============================================
echo   Amazon 评论爬虫 - 打包工具
echo ============================================
echo.

:: 检查 Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到 Python，请先安装 Python 3.8+
    pause
    exit /b 1
)

:: 安装打包依赖
echo [1/2] 安装 PyInstaller ...
pip install pyinstaller playwright >nul 2>&1

:: 打包（不再内置浏览器，运行时自动使用系统 Chrome）
echo [2/2] 正在打包为独立可执行文件 ...
pyinstaller ^
    --onedir ^
    --name AmazonReviewScraper ^
    --noconsole ^
    --noconfirm ^
    --clean ^
    --hidden-import playwright ^
    --hidden-import playwright.sync_api ^
    --hidden-import greenlet ^
    --hidden-import openpyxl ^
    --hidden-import openpyxl.styles ^
    --hidden-import openpyxl.utils ^
    --collect-submodules openpyxl ^
    --exclude-module numpy ^
    --exclude-module pandas ^
    --exclude-module numba ^
    --exclude-module llvmlite ^
    --exclude-module PIL ^
    --exclude-module Pillow ^
    --exclude-module scipy ^
    --exclude-module matplotlib ^
    --exclude-module cryptography ^
    --exclude-module win32com ^
    --exclude-module Pythonwin ^
    --exclude-module setuptools ^
    --exclude-module tqdm ^
    amazon_review_scraper.py

:: 创建启动脚本
echo @echo off > dist\AmazonReviewScraper\run.bat
echo chcp 65001 ^>nul >> dist\AmazonReviewScraper\run.bat
echo "%%~dp0AmazonReviewScraper.exe" >> dist\AmazonReviewScraper\run.bat
echo pause >> dist\AmazonReviewScraper\run.bat

echo.
echo ============================================
echo   打包完成！
echo ============================================
echo.
echo 输出目录: dist\AmazonReviewScraper\
echo 运行方式: 双击 dist\AmazonReviewScraper\run.bat
echo.
echo 将整个 dist\AmazonReviewScraper 文件夹拷贝给他人即可直接使用。
echo 目标机器需已安装 Google Chrome。
echo.
pause

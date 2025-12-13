@echo off
:: 强制切换为GBK编码（CMD默认编码，彻底解决中文乱码，比UTF-8更兼容）
chcp 936 > nul
:: 启用延迟扩展
setlocal enabledelayedexpansion

:: 动态获取当前Git仓库的根目录路径（关键优化：通用化）
for /f "tokens=*" %%i in ('git rev-parse --show-toplevel 2^>nul') do (
    set "repo_path=%%i"
)

:: 设置窗口标题（使用动态仓库路径）
title 一键上传代码（默认提交：作者修改）- !repo_path!
echo ==============================================
echo          一键上传代码到GitHub脚本
echo          （默认提交说明：作者修改）
echo          （当前仓库：!repo_path!）
echo ==============================================
echo.

:: 步骤1：检查是否在Git仓库根目录
git rev-parse --is-inside-work-tree > nul 2>&1
if errorlevel 1 (
    echo 错误：脚本不在Git仓库根目录运行！
    pause > nul
    exit /b 1
)

:: 步骤2：同步本地所有变更（新增/修改/删除）到暂存区
echo 1. 正在同步本地所有变更到暂存区...
git add -A
if errorlevel 1 (
    echo 错误：同步变更失败，请检查文件是否被占用！
    pause > nul
    exit /b 1
)

:: 步骤3：检测是否有变更（避免无变更时提交）
set "has_change=0"
for /f "tokens=*" %%i in ('git status --porcelain') do (
    set "has_change=1"
)
if !has_change! equ 0 (
    echo 提示：本地文件无任何变更，无需提交和推送。
    pause > nul
    exit /b 0
)

:: 步骤4：可选自定义提交说明（保留默认，也可手动输入）
set "commit_msg=作者修改"
echo.
echo 请输入提交说明（直接回车使用默认：%commit_msg%）：
set /p "input_msg="
:: 如果用户输入了内容，就替换默认提交说明
if not "!input_msg!"=="" (
    set "commit_msg=!input_msg!"
)
echo 2. 正在提交到本地仓库（说明：!commit_msg!）...
git commit -m "!commit_msg!"
if errorlevel 1 (
    echo 错误：提交失败，请手动执行git status查看原因！
    pause > nul
    exit /b 1
)

:: 步骤5：推送到GitHub云端仓库
echo 3. 正在推送到GitHub远程仓库...
git push
if errorlevel 1 (
    echo 错误：推送失败，请检查网络或GitHub登录状态！
    pause > nul
    exit /b 1
)

echo.
echo ==============================================
echo 上传成功！云端仓库已与本地完全一致！
echo 提交说明：!commit_msg!
echo ==============================================
pause > nul
exit /b 0
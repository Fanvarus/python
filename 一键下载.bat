@echo off
:: 切换为CMD默认的GBK编码（解决中文乱码）
chcp 936 > nul
:: 启用延迟扩展
setlocal enabledelayedexpansion

:: 动态获取当前Git仓库的根目录路径（替代硬编码路径）
for /f "tokens=*" %%i in ('git rev-parse --show-toplevel 2^>nul') do (
    set "repo_path=%%i"
)

:: 动态获取当前本地分支名称（解决分支固定为main的问题）
for /f "tokens=*" %%i in ('git rev-parse --abbrev-ref HEAD 2^>nul') do (
    set "current_branch=%%i"
)

:: 设置窗口标题（加入动态仓库路径和分支）
title 一键下载GitHub代码（云端强制覆盖本地）- !repo_path!
echo ==============================================
echo          一键下载GitHub最新代码脚本
echo          （云端强制覆盖本地所有文件）
echo          （当前仓库：!repo_path!）
echo          （当前分支：!current_branch!）
echo ==============================================
echo.

:: 步骤1：检查是否在Git仓库根目录
git rev-parse --is-inside-work-tree > nul 2>&1
if errorlevel 1 (
    echo 错误：脚本不在Git仓库根目录运行！
    pause > nul
    exit /b 1
)

:: 步骤2：提示风险（强制覆盖会丢失本地未提交修改）
echo 注意：此操作会用云端%current_branch%分支的最新代码强制覆盖本地！
echo 本地未提交的修改将被永久丢失，请确认后按任意键继续...
pause > nul
echo.

:: 步骤3：拉取云端最新代码到本地缓存（使用当前分支）
echo 1. 正在从GitHub拉取%current_branch%分支最新代码缓存...
git fetch origin !current_branch!
if errorlevel 1 (
    echo 错误：拉取代码缓存失败，请检查网络或GitHub地址！
    echo 可能原因：云端无%current_branch%分支，或仓库未关联远程！
    pause > nul
    exit /b 1
)

:: 步骤4：强制将本地分支重置为云端对应分支（覆盖本地）
echo 2. 正在用云端%current_branch%分支代码强制覆盖本地文件...
git reset --hard origin/!current_branch!
if errorlevel 1 (
    echo 错误：强制覆盖本地文件失败！
    pause > nul
    exit /b 1
)

echo.
echo ==============================================
echo 下载成功！本地文件已被云端%current_branch%分支强制覆盖！
echo ==============================================
pause > nul
exit /b 0
@echo off
echo 正在自动同步中...
cd /d "E:\天之逸2025\Moze4.0"

:: 先把云端的拉下来（防冲突）
git pull

:: 添加所有修改
git add .

:: 自动提交（以时间作为备注）
git commit -m "Auto-sync %date% %time%"

:: 推送到所有平台
git push

echo.
echo 搞定！3秒后自动关闭...
timeout /t 3
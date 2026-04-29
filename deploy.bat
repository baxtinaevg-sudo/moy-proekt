@echo off
cd /d c:\Users\Нюся\additiv-site
rmdir /s /q .git 2>nul
git init
git remote add origin https://git.msk0.amvera.ru/evgeniia/additiv
git add -A
git commit -m "Сайт Аддитив Плюс"
git branch -M main
git push -f origin main
pause
@echo off
cd /d c:\Users\Нюся\additiv-site
git add -A
git commit -m "Сайт Аддитив Плюс"
git branch -M main
git push -f github main
pause
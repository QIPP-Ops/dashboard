@echo off
cd /d "C:\Users\asus\WebstormProjects\QIPP-Dashboard"
git add .
git commit -m "Auto-update %date% %time%"
git push origin main

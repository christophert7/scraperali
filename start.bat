@echo off
rem ===== AliExpress Product Scraper =====
rem انقر نقرة مزدوجة على هذا الملف لتشغيل الواجهة الرسومية
cd /d "%~dp0"
python gui.py
if errorlevel 1 pause

@echo off
chcp 65001 > nul
title Tool Kẻ Dòng Đề Thi Word
echo Đang tự động xử lý các file trong thư mục input...
py main.py
echo.
echo Đã hoàn tất! Nhấn phím bất kỳ để thoát...
pause > nul
@echo off
title ショートカット作成 (文書横断検索)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0make_shortcut.ps1"
echo.
pause

@echo off
setlocal

REM Launch HWPX -> HTML GUI converter
REM Requirements: Python installed, deps installed, HWP installed (Windows)

cd /d "%~dp0"
python -X utf8 "%~dp0convert_hwpx_to_html_gui.py"

endlocal


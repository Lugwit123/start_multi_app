@echo off
setlocal

REM Resolve package root: ...\src\start_multi_app\bat -> ...\ (package root)
set "BAT_DIR=%~dp0"
for %%I in ("%BAT_DIR%..\..\..\") do set "PKG_ROOT=%%~fI"
set "MAIN_PY=%PKG_ROOT%src\start_multi_app\main.py"

if not exist "%MAIN_PY%" (
    echo [start_multi_app] main.py not found: "%MAIN_PY%"
    exit /b 1
)

python "%MAIN_PY%" %*
exit /b %errorlevel%

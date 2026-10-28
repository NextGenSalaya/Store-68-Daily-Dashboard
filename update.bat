@echo off
rem Daily update of the Store 68 Daily Board website.
rem 1) put today's Excel files in this folder (sub-folders are fine)  2) double-click this file
chcp 65001 >nul
cd /d "%~dp0"
rem Git for Windows installed per user is not always on PATH
where git >nul 2>&1 || set "PATH=%LOCALAPPDATA%\Programs\Git\cmd;%PATH%"
echo === Reading the Excel files ===
python build_all_dashboard.py || goto :fail
python build_topic2.py || goto :fail
python build_topic3.py || goto :fail
python build_topic4.py || goto :fail
python build_topic5.py || goto :fail
echo === Building the website ===
python make_site.py || goto :fail
echo === Publishing to GitHub ===
git add -A
git commit -m "Update data %date%"
git push || goto :fail
echo.
echo Done. The website updates in about a minute.
pause
exit /b 0
:fail
echo.
echo Something went wrong - see the message above. Nothing was published after the failed step.
pause
exit /b 1

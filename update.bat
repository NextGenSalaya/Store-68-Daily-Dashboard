@echo off
rem Daily update of the Store 68 Daily Board website.
rem 1) put today's Excel files in this folder (sub-folders are fine)  2) double-click this file
rem 3) drag everything from the publish folder that opens onto the GitHub page that opens, then press Commit changes
chcp 65001 >nul
cd /d "%~dp0"
echo === Reading the Excel files ===
python build_all_dashboard.py || goto :fail
python build_topic2.py || goto :fail
python build_topic3.py || goto :fail
python build_topic4.py || goto :fail
python build_topic5.py || goto :fail
echo === Building the website ===
python make_site.py || goto :fail
echo.
echo Done. Next:
echo   1. In the "publish" folder that opens, press Ctrl+A and drag everything onto the GitHub page.
echo      (only index.html and .json files - never drag .xlsx files)
echo   2. Wait until every file has uploaded, then press "Commit changes".
echo   3. About 1 minute later, refresh https://nextgensalaya.github.io/Store-68-Daily-Dashboard/
start "" explorer "%~dp0publish"
start "" "https://github.com/NextGenSalaya/Store-68-Daily-Dashboard/upload/main"
pause
exit /b 0
:fail
echo.
echo Something went wrong - see the message above. Nothing was published.
pause
exit /b 1

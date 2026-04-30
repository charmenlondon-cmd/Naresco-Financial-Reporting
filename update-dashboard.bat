@echo off
echo ============================================================
echo NARESCO DASHBOARD UPDATE SCRIPT
echo ============================================================
echo.

REM Change to script directory
cd /d "%~dp0"

echo [1/5] Finding newest Excel file in source-files folder...
echo.

REM Use PowerShell to find the newest .xlsx file
for /f "delims=" %%i in ('powershell -Command "Get-ChildItem -Path 'source-files\*.xlsx' | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty Name"') do set NEWEST_FILE=%%i

if "%NEWEST_FILE%"=="" (
    echo ERROR: No Excel files found in source-files folder!
    echo Please add your Budget vs Actual Excel file to the source-files folder.
    pause
    exit /b 1
)

echo Found: %NEWEST_FILE%
echo.

echo [2/5] Extracting data from Excel file...
python scripts\extract_budget_data.py --input "source-files\%NEWEST_FILE%"
if %ERRORLEVEL% neq 0 (
    echo ERROR: Data extraction failed!
    pause
    exit /b 1
)
echo.

echo [3/5] Consolidating monthly data to YTD totals...
python scripts\consolidate_data.py
if %ERRORLEVEL% neq 0 (
    echo ERROR: Data consolidation failed!
    pause
    exit /b 1
)
echo.

echo [4/5] Generating dashboard HTML...
python scripts\generate_dashboard.py
if %ERRORLEVEL% neq 0 (
    echo ERROR: Dashboard generation failed!
    pause
    exit /b 1
)
echo.

echo [5/5] Deploying to web (git commit and push)...
git add dashboard\index.html
git commit -m "Update dashboard with latest data from %NEWEST_FILE%"
git push
if %ERRORLEVEL% neq 0 (
    echo WARNING: Git push failed. Check your git credentials.
    pause
    exit /b 1
)
echo.

echo ============================================================
echo SUCCESS! Dashboard updated and deployed!
echo ============================================================
echo.
echo Your dashboard will be live at:
echo https://naresco-financial-reporting.vercel.app
echo.
echo (Vercel deployment takes ~30 seconds)
echo.
pause

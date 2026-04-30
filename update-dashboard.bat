@echo off
echo ============================================================
echo NARESCO DASHBOARD UPDATE SCRIPT
echo ============================================================
echo.

REM Change to script directory
cd /d "%~dp0"

echo [1/8] Finding newest Excel file in source-files folder...
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

echo [2/8] PATH A: Extracting to Excel Database...
python scripts\extract_to_excel_database.py --input "source-files\%NEWEST_FILE%"
if %ERRORLEVEL% neq 0 (
    echo ERROR: Excel database extraction failed!
    pause
    exit /b 1
)
echo.

echo [3/8] PATH B: Extracting to JSON...
python scripts\extract_budget_data.py --input "source-files\%NEWEST_FILE%"
if %ERRORLEVEL% neq 0 (
    echo ERROR: JSON extraction failed!
    pause
    exit /b 1
)
echo.

echo [4/8] PATH B: Consolidating JSON data to YTD totals...
python scripts\consolidate_data.py
if %ERRORLEVEL% neq 0 (
    echo ERROR: JSON consolidation failed!
    pause
    exit /b 1
)
echo.

echo [5/8] VALIDATION: Comparing Excel vs JSON calculations...
python scripts\compare_calculations.py
if %ERRORLEVEL% neq 0 (
    echo.
    echo ============================================================
    echo VALIDATION FAILED!
    echo Excel and JSON calculations do not match.
    echo Please investigate before deploying.
    echo ============================================================
    pause
    exit /b 1
)
echo.

echo [6/8] Generating dashboard from Excel (validated source)...
python scripts\generate_dashboard_from_excel.py
if %ERRORLEVEL% neq 0 (
    echo ERROR: Dashboard generation failed!
    pause
    exit /b 1
)
echo.

echo [7/8] Deploying to web (git commit and push)...
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
echo Data Source: Excel Database (validated against JSON)
echo All calculations verified and matched!
echo.
echo Your dashboard will be live at:
echo https://naresco-financial-reporting.vercel.app
echo.
echo (Vercel deployment takes ~30 seconds)
echo.
pause

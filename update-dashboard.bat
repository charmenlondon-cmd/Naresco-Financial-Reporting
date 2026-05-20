@echo off
echo ============================================================
echo NARESCO DASHBOARD UPDATE SCRIPT (Multi-Company)
echo ============================================================
echo.

cd /d "%~dp0"

REM Process all Excel files in source-files folder
set PROCESSED=0
set FAILED=0

for /f "delims=" %%f in ('dir /b /a-d "source-files\*.xlsx" 2^>nul') do (
    call :PROCESS_FILE "%%f"
)

if %PROCESSED%==0 (
    echo ERROR: No Excel files found in source-files folder!
    echo Please add Budget vs Actual Excel files to the source-files folder.
    pause
    exit /b 1
)

echo.
echo ============================================================
if %FAILED%==0 (
    echo SUCCESS! Processed %PROCESSED% company file(s^).
) else (
    echo COMPLETED with errors: %PROCESSED% processed, %FAILED% failed.
)
echo ============================================================
echo.
pause
exit /b 0

:PROCESS_FILE
set FILENAME=%~1
echo.
echo ------------------------------------------------------------
echo Processing: %FILENAME%
echo ------------------------------------------------------------

REM Detect company from filename
set COMPANY=
echo %FILENAME% | findstr /i "mudin mae" >nul && set COMPANY=mudin
if "%COMPANY%"=="" echo %FILENAME% | findstr /i "mantis" >nul && set COMPANY=mantis

if "%COMPANY%"=="" (
    echo WARNING: Could not detect company from filename "%FILENAME%". Skipping.
    set /a FAILED+=1
    exit /b 0
)

echo Detected company: %COMPANY%

REM Detect file type from filename
set FILETYPE=
echo %FILENAME% | findstr /i "Financial Statement" >nul && set FILETYPE=fs
if "%FILETYPE%"=="" set FILETYPE=bva

echo Detected file type: %FILETYPE%
echo.

if "%FILETYPE%"=="fs" goto :FS_PIPELINE

REM -----------------------------------------------
REM BVA PIPELINE (Budget vs Actual file)
REM -----------------------------------------------
echo [1/6] Extracting to Excel Database...
python scripts\extract_to_excel_database.py --input "source-files\%FILENAME%" --company %COMPANY%
if %ERRORLEVEL% neq 0 ( echo ERROR: Extraction to database failed! && set /a FAILED+=1 && exit /b 1 )

echo [2/6] Extracting section detail data (Variable Costs etc.)...
python scripts\extract_section_data.py --input "source-files\%FILENAME%" --company %COMPANY%
if %ERRORLEVEL% neq 0 ( echo WARNING: Section extraction had errors - detail drill-downs may be incomplete )

echo [3/6] Extracting to JSON...
python scripts\extract_budget_data.py --input "source-files\%FILENAME%" --company %COMPANY%
if %ERRORLEVEL% neq 0 ( echo ERROR: JSON extraction failed! && set /a FAILED+=1 && exit /b 1 )

echo [4/6] Consolidating JSON data to YTD totals...
python scripts\consolidate_data.py --company %COMPANY%
if %ERRORLEVEL% neq 0 ( echo ERROR: Consolidation failed! && set /a FAILED+=1 && exit /b 1 )

echo [5/6] Validating Excel vs JSON calculations...
python scripts\compare_calculations.py --company %COMPANY%
if %ERRORLEVEL% neq 0 ( echo VALIDATION FAILED for %COMPANY%! && set /a FAILED+=1 && exit /b 1 )

echo [6/6] Generating dashboard data...
python scripts\generate_dashboard_from_excel.py --company %COMPANY%
if %ERRORLEVEL% neq 0 ( echo ERROR: Dashboard data generation failed! && set /a FAILED+=1 && exit /b 1 )

set /a PROCESSED+=1
echo [OK] %COMPANY% BVA pipeline completed successfully.
exit /b 0

REM -----------------------------------------------
:FS_PIPELINE
REM FINANCIAL STATEMENT PIPELINE (revenue line items)
REM -----------------------------------------------
echo [1/1] Extracting revenue data to database...
python scripts\extract_revenue_data.py --input "source-files\%FILENAME%" --company %COMPANY%
if %ERRORLEVEL% neq 0 ( echo ERROR: Revenue extraction failed! && set /a FAILED+=1 && exit /b 1 )

set /a PROCESSED+=1
echo [OK] %COMPANY% revenue extraction completed successfully.
exit /b 0

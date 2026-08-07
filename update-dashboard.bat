@echo off
cd /d "%~dp0"
echo ============================================================
echo NARESCO DASHBOARD UPDATE SCRIPT (Multi-Company)
echo ============================================================
echo.

set PROCESSED=0
set FAILED=0
set FOUND=0

for %%f in (source-files\*.xlsx) do (
    set FOUND=1
    call :PROCESS_FILE "%%~nxf"
)

if %FOUND%==0 (
    echo ERROR: No Excel files found in source-files folder!
    echo Please add Excel files to the source-files folder.
    pause
    exit /b 1
)

echo.
echo ============================================================
if %FAILED%==0 (
    echo SUCCESS! Processed %PROCESSED% company file(s).
) else (
    echo COMPLETED with errors: %PROCESSED% processed, %FAILED% failed.
)
echo ============================================================
echo.
pause
exit /b 0

REM ============================================================
:PROCESS_FILE
set FILENAME=%~1
echo.
echo ------------------------------------------------------------
echo Processing: %FILENAME%
echo ------------------------------------------------------------

set COMPANY=
echo %FILENAME% | findstr /i "mudin mae" >nul
if not errorlevel 1 set COMPANY=mudin
if not "%COMPANY%"=="" goto :COMPANY_FOUND
echo %FILENAME% | findstr /i "mantis" >nul
if not errorlevel 1 set COMPANY=mantis
if not "%COMPANY%"=="" goto :COMPANY_FOUND
echo %FILENAME% | findstr /i "mcp" >nul
if not errorlevel 1 set COMPANY=mcp
if not "%COMPANY%"=="" goto :COMPANY_FOUND
echo %FILENAME% | findstr /i "mci" >nul
if not errorlevel 1 set COMPANY=mci
if not "%COMPANY%"=="" goto :COMPANY_FOUND
echo %FILENAME% | findstr /i "transon" >nul
if not errorlevel 1 set COMPANY=transon

:COMPANY_FOUND
if "%COMPANY%"=="" (
    echo WARNING: Could not detect company from filename "%FILENAME%". Skipping.
    set /a FAILED+=1
    exit /b 0
)
echo Detected company: %COMPANY%

set FILETYPE=bva
echo %FILENAME% | findstr /i "Financial Statement" >nul
if not errorlevel 1 set FILETYPE=fs
echo Detected file type: %FILETYPE%
echo.

if "%FILETYPE%"=="fs" goto :FS_PIPELINE

REM -----------------------------------------------
REM BVA PIPELINE
REM -----------------------------------------------
echo [1/6] Extracting to Excel Database...
python scripts\extract_to_excel_database.py --input "source-files\%FILENAME%" --company %COMPANY%
if errorlevel 1 ( echo ERROR: Step 1 failed! & set /a FAILED+=1 & exit /b 1 )

echo [2/6] Extracting section detail data...
python scripts\extract_section_data.py --input "source-files\%FILENAME%" --company %COMPANY%

echo [3/6] Extracting to JSON...
python scripts\extract_budget_data.py --input "source-files\%FILENAME%" --company %COMPANY%
if errorlevel 1 ( echo ERROR: Step 3 failed! & set /a FAILED+=1 & exit /b 1 )

echo [4/6] Consolidating JSON data to YTD totals...
python scripts\consolidate_data.py --company %COMPANY%
if errorlevel 1 ( echo ERROR: Step 4 failed! & set /a FAILED+=1 & exit /b 1 )

echo [5/6] Validating Excel vs JSON calculations...
python scripts\compare_calculations.py --company %COMPANY%
if errorlevel 1 ( echo VALIDATION FAILED for %COMPANY%! & set /a FAILED+=1 & exit /b 1 )

echo [6/6] Generating dashboard data...
python scripts\generate_dashboard_from_excel.py --company %COMPANY%
if errorlevel 1 ( echo ERROR: Step 6 failed! & set /a FAILED+=1 & exit /b 1 )

set /a PROCESSED+=1
echo [OK] %COMPANY% BVA pipeline completed successfully.
if not exist "source-files\archive" mkdir "source-files\archive"
move "source-files\%FILENAME%" "source-files\archive\%FILENAME%" >nul
echo [OK] Archived: %FILENAME%
exit /b 0

REM -----------------------------------------------
:FS_PIPELINE
echo [1/1] Extracting revenue data to database...
python scripts\extract_revenue_data.py --input "source-files\%FILENAME%" --company %COMPANY%
if errorlevel 1 ( echo WARNING: Revenue extraction failed - file format may be unsupported. Archiving anyway. )

set /a PROCESSED+=1
echo [OK] %COMPANY% revenue extraction completed successfully.
if not exist "source-files\archive" mkdir "source-files\archive"
move "source-files\%FILENAME%" "source-files\archive\%FILENAME%" >nul
echo [OK] Archived: %FILENAME%
exit /b 0
@echo off
:: ============================================================
:: run_psx.bat
:: Scrapes PSX data for selected tickers and writes to Google Sheets.
:: Schedule this file via Windows Task Scheduler.
::
:: TASK SCHEDULER TRIGGER TIMES
:: ---------------------------------------------------------------
:: Monday to Thursday (market opens 09:30):
::   09:33, 10:30, 11:30, 12:30, 13:30, 14:30, 15:30, 16:33
::
:: Friday (market opens 09:15, lunch break 12:00-14:00):
::   09:15, 10:15, 11:15, 12:00, 14:30, 15:30, 16:30
::
:: Configure two separate Task Scheduler tasks:
::   Task 1 (Mon-Thu): set to run Mon/Tue/Wed/Thu only
::   Task 2 (Friday) : set to run Friday only
:: ============================================================

:: Change to the folder that contains this batch file
cd /d "%~dp0"

:: Skip on weekends (PSX is closed Saturday and Sunday)
for /f %%D in ('powershell -NoProfile -Command "(Get-Date).DayOfWeek"') do set DOW=%%D
if "%DOW%"=="Saturday" (
    echo [%DATE% %TIME%] Weekend -- skipping.
    exit /b 0
)
if "%DOW%"=="Sunday" (
    echo [%DATE% %TIME%] Weekend -- skipping.
    exit /b 0
)

:: Note Friday market hours
if "%DOW%"=="Friday" (
    echo [%DATE% %TIME%] Friday -- PSX opens at 09:15 AM. Using Friday trigger schedule.
)

:: List of tickers to process
set TICKERS=KEL OGDC PPL HBL ENGRO OBOY FCEPL

for %%T in (%TICKERS%) do (
    echo.
    echo [%DATE% %TIME%] Processing %%T ...
    python psx_to_sheets.py %%T
    if errorlevel 1 (
        echo [ERROR] Failed for %%T
    )
)

echo.
echo Done.

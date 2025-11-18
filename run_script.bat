@echo off
chcp 65001 >nul
cd /d "%~dp0"

set COLOR_SUCCESS=0A
set COLOR_ERROR=0C
set COLOR_INFO=0B

if not exist "venv\" (
    color %COLOR_INFO%
    echo *** Creating virtual environment... ***
    py -m venv venv
)

call venv\Scripts\activate.bat

for /f "tokens=2 delims==" %%I in ('wmic os get localdatetime /value') do set dt=%%I
set CURRENT_DATE=%dt:~0,4%-%dt:~4,2%-%dt:~6,2%
set LOG_DATE=%dt:~0,4%-%dt:~4,2%-%dt:~6,2%_%dt:~8,2%-%dt:~10,2%-%dt:~12,2%

if not exist "logs" mkdir logs
set LOGFILE=%~dp0logs\log_%LOG_DATE%.txt

set /a count=0
color %COLOR_INFO%

echo *** Choose one of the following variants: ***
echo *** 1. Make a report of the system active users ***
echo *** 2. Block the inactive users ***
set /p choice=Enter your choice (1 or 2):

if "%choice%"=="1" goto run_report
if "%choice%"=="2" goto block_users

echo Invalid choice!
exit /b 1

:: ----------------   REPORT   ----------------
:run_report
echo *** Running report_creator.py script... ***
call py report_creator.py --user-report > "%LOGFILE%" 2>&1

find /i "Error" "%LOGFILE%" >nul
if %errorlevel%==0 (
    color %COLOR_ERROR%
    echo *** Errors occurred. See %LOGFILE% ***

    find /i "No module" "%LOGFILE%" >nul
    if %errorlevel%==0 (
        if %count%==0 (
            if exist "requirements.txt" (
                color %COLOR_INFO%
                echo *** Installing dependencies... ***
                py -m pip install --upgrade pip
                pip install -r requirements.txt >nul
                set /a count+=1
                goto run_report
            )
        )
    )
    exit /b 1
)

if not exist "report_creator.py" (
    color %COLOR_ERROR%
    echo *** File report_creator.py not found! ***
    exit /b 1
)

color %COLOR_SUCCESS%
del "%LOGFILE%"
echo *** File "Безпечне місто_%CURRENT_DATE%.xlsx" was created! ***
echo See the reports/ directory.
goto end


:: ----------------   BLOCKING   ----------------
:block_users
echo *** Running report_creator.py script... ***
call py report_creator.py --block-inactive > "%LOGFILE%" 2>&1

find /i "Error" "%LOGFILE%" >nul
if %errorlevel%==0 (
    color %COLOR_ERROR%
    echo *** Errors occurred. See %LOGFILE% ***

    find /i "No module" "%LOGFILE%" >nul
    if %errorlevel%==0 (
        if %count%==0 (
            if exist "requirements.txt" (
                color %COLOR_INFO%
                echo *** Installing dependencies... ***
                py -m pip install --upgrade pip
                pip install -r requirements.txt >nul
                set /a count+=1
                goto block_users
            )
        )
    )
    exit /b 1
)

set FILEPATH=%~dp0blocked\Blocked inactive users %CURRENT_DATE%.xlsx

:: Перевірка на пустоту файла
if exist "%FILEPATH%" (
    for %%A in ("%FILEPATH%") do if %%~zA==0 (
        del "%FILEPATH%"
        color %COLOR_ERROR%
        echo *** File was empty and deleted. ***
        goto end
    )
)

color %COLOR_SUCCESS%
del "%LOGFILE%"
echo *** File "Blocked inactive users %CURRENT_DATE%.xlsx" was created! ***
echo *** See blocked/ directory with the blocked users report file. ***

:end
pause

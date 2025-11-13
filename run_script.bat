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
:restart_label
color %COLOR_INFO%
echo *** Running report_creator.py script... ***
py "report_creator.py" > "%LOGFILE%" 2>&1

find /i "Error" "%LOGFILE%" >nul
if %errorlevel%==0 (
    color %COLOR_ERROR%
    echo *** Script running errors occurred. See %LOGFILE% ***
	find /i "No module" "%LOGFILE%" >nul
	if %errorlevel%==0 (
		if %count%==0 (
			if exist "requirements.txt" (
				color %COLOR_INFO%
				echo *** Installing dependencies... ***
				py -m pip install --upgrade pip
				pip install -r requirements.txt >nul
				set /a count+=1
				echo *** Re-running report_creator.py script... ***
				goto restart_label
			)
		)
	)
) else (
    if not exist "report_creator.py" (
        color %COLOR_ERROR%
        echo *** File report_creator.py not found! ***
    ) else (
        color %COLOR_SUCCESS%
		del %LOGFILE%
		echo. 
        echo *** File "Безпечне місто_%CURRENT_DATE%.xlsx" was created!!! See the reports/ directory. ***
    )
)

pause

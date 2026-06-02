@echo off
REM Windows Task Scheduler wrapper for fetch_qri_iv.py
REM Schedule: run every 15 minutes during NK225 option trading hours.
REM
REM Manual one-shot:   scripts\fetch_qri_iv.bat
REM Loop (foreground): scripts\fetch_qri_iv.bat --loop
REM
REM To register as a 15-min scheduled task (open as Administrator):
REM   schtasks /Create /SC MINUTE /MO 15 /TN "QRI_IV_Fetch" ^
REM     /TR "C:\Users\kawai\Desktop\実験室01\先物手口分析\scripts\fetch_qri_iv.bat" ^
REM     /ST 08:45 /F

cd /d "%~dp0.."
python scripts\fetch_qri_iv.py %*
exit /b %ERRORLEVEL%

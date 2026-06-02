@echo off
REM Windows Task Scheduler wrapper for fetch_qri_iv.py
REM Uploads to R2 by default. Pass extra flags through (%*).
REM
REM Manual one-shot (local + R2): scripts\fetch_qri_iv.bat
REM Loop (foreground):            scripts\fetch_qri_iv.bat --loop
REM R2 only (no local file):      scripts\fetch_qri_iv.bat --no-local
REM
REM To register as a 15-min scheduled task (open as Administrator):
REM   schtasks /Create /SC MINUTE /MO 15 /TN "QRI_IV_Fetch" ^
REM     /TR "C:\Users\kawai\Desktop\実験室01\先物手口分析\scripts\fetch_qri_iv.bat" ^
REM     /ST 08:45 /F

cd /d "%~dp0.."
python scripts\fetch_qri_iv.py --r2 %*
exit /b %ERRORLEVEL%

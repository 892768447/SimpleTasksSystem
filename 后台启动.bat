cd %~dp0
taskkill /F /IM pythonw.exe
start /b pythonw AutoReportMms.py >> datas/error.log
pause
@echo off
powershell -windowstyle normal -command "$s=(New-Object -ComObject Shell.Application).NameSpace('%~dp0').ParseName('PuliziaUltraSafe.ico');"

powershell.exe -ExecutionPolicy Bypass -File "%~dp0Antho_Tool.ps1"
exit

@echo off
set SCRIPT=%~dp0PuliziaUltraSafe.ps1
powershell -ExecutionPolicy Bypass -File "%SCRIPT%"
pause

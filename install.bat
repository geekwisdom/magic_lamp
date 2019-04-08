@echo off
set install=%~dp0
Echo Current Directory is %install%
schtasks /Create /RU %username% /SC ONEVENT /MO "*[System[Provider[@Name='Microsoft-Windows-Security-Auditing'] and EventID=4801]]" /EC Security /TN "Motivational Quote" /TR "%windir%\system32\wscript.exe %install%\inspiration.vbs" /F
Echo Install Completed
pause


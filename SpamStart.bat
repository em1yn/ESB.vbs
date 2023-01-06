@echo off
set loops=20
:loop
start EmailSpam.vbs
TIMEOUT /T 300
taskkill /im wscript.exe /f 
set /a loopcount=loopcount-1
if %loops%==0 goto exitloop
goto loop
:exitloop

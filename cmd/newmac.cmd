@ECHO OFF

REM Set Variable of Location of "macshift.exe"
SET BINDIR=C:\bin

REM Create temporary helper batch file to
REM store hundredths of seconds in variable
echo set random=%%9>>%temp%\enter.bat

REM Store current time in temporary batch file
ver | time | date | find "Current" | find ")" > %temp%\temp.bat

REM Store current time's hundredths of seconds in variable
call %temp%\temp.bat

REM Remove helper files
del %temp%\temp.bat
del %temp%\enter.bat

REM Display result
REM echo Random number: %random%

%BINDIR%\macshift -i "Wireless Network Connection 3" 0001F4EE%random%
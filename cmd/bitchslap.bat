@ECHO OFF
SET COMP=%1
SET MSG=%2
net send %COMP% %MSG%
call bitchslap.bat %COMP% %MSG%

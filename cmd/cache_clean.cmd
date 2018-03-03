:: -------------------------------------------------
:: -- Cache Cleaner for Windows 2000/XP/2003
:: --
:: -- FileName: cache_clean.cmd
:: -- Revision: 2.1
:: -- Created: 19-AUG-2005
:: -- Author: Mike Pruett
:: -- Disclaimer: "Review each delete statement as this file deletes numerous caches.
:: -- 		 File to automate deletion of Jinitiator cache and temp files when
:: --		 a user clicks desktop icon."
:: -------------------------------------------------
@ECHO OFF
echo @ECHO OFF > "c:\ccleaner.cmd"
:: -- JInitiator Cache Cleaning Section
for /f "tokens=1,2 delims= " %%A in ('DIR /AD /B "%PROGRAMFILES%\Oracle" ^| find "JInitiator"') do (set J_CACHE=%%A& set J_VER=%%B)
echo del "%PROGRAMFILES%\Oracle\%J_CACHE% %J_VER%\Jcache\*.*" /q >> "c:\ccleaner.cmd"

:: -- User Temp Cleaning Section
echo rd "%USERPROFILE%\Local Settings\Temp" /s /q >> "c:\ccleaner.cmd"
echo del "%USERPROFILE%\Local Settings\Temp\*.*" /s /q >> "c:\ccleaner.cmd"
echo mkdir "%USERPROFILE%\Local Settings\Temp" >> "c:\ccleaner.cmd"
echo rd "%WINDIR%\Temp" /s /q >> "c:\ccleaner.cmd"
echo rd "%WINDIR%\Temp\*.*" /s /q >> "c:\ccleaner.cmd"
echo mkdir "%WINDIR%\Temp" >> "c:\ccleaner.cmd"

:: -- Netscape Cache Cleaning Section
echo del "%SystemDrive%\Program Files\Netscape\Users\default\Cache\*.*" /q >> "c:\ccleaner.cmd"

:: -- Internet Explorer Cache Cleaning Section
echo rd "%USERPROFILE%\Local Settings\Temporary Internet Files\" /q /s >> "c:\ccleaner.cmd"
echo mkdir "%USERPROFILE%\Local Settings\Temporary Internet Files" >> "c:\ccleaner.cmd"

:: -- Firefox Cache Cleaning Section
for /f "tokens=1" %%A in ('DIR /AD /B "%APPDATA%\Mozilla\Firefox\Profiles"') do (set MOZPROF=%%A)
echo del "%APPDATA%\Mozilla\Firefox\Profiles\%MOZPROF%\Cache\*.*" /q >> "c:\ccleaner.cmd"
echo del "%USERPROFILE%\ccleaner.reg" /q /s >> "c:\ccleaner.cmd"
echo del "c:\ccleaner.cmd" /q /s >> "c:\ccleaner.cmd"
echo exit >> "c:\ccleaner.cmd"

:: -- Create the Registry File to import into the RunOnce Key
echo Windows Registry Editor Version 5.00 > "%USERPROFILE%\ccleaner.reg"
echo [HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce] >> "%USERPROFILE%\ccleaner.reg"
echo "RemoveCache"="C:\\ccleaner.cmd" >> "%USERPROFILE%\ccleaner.reg"

:: -- Add ccleaner.reg to the Registry
regedit /s "%USERPROFILE%\ccleaner.reg"

::End
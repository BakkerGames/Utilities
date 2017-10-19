@echo off

attrib -r Build*.bat >nul 2>nul
del Build*.bat       >nul 2>nul

set UPDVERS="P:\Bin\Utilities\UpdateVersions2.exe"
if not exist %UPDVERS% set UPDVERS="\Bin\Utilities\UpdateVersions2.exe"

%UPDVERS% "."
@if errorlevel 1 goto error

call BuildAll.bat
@if errorlevel 1 goto error

del BuildAll.bat>nul 2>nul
goto done

:error
pause
:done

@echo off

attrib -r Build*.bat >nul 2>nul
del Build*.bat       >nul 2>nul

set UPDVERS="P:\Utilities\Bin\UpdateVersions2.exe"
if not exist %UPDVERS% set UPDVERS="C:\Utilities\Bin\UpdateVersions2.exe"
if not exist %UPDVERS% set UPDVERS="\Utilities\Bin\UpdateVersions2.exe"

%UPDVERS% "."

pause

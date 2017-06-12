@echo off

attrib -r Build*.bat >nul 2>nul
del Build*.bat       >nul 2>nul

set UPDVERS="P:\Utilities\Bin\UpdateVersions2.exe"
if not exist %UPDVERS% set UPDVERS="C:\Utilities\Bin\UpdateVersions2.exe"
if not exist %UPDVERS% set UPDVERS="\Utilities\Bin\UpdateVersions2.exe"

%UPDVERS% "."

attrib -r /s Bin\*.* >nul 2>nul

call BuildAll.bat

del BuildAll.bat>nul 2>nul

attrib +r /s Bin\*.* >nul 2>nul

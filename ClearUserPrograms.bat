@echo off
SET TARGETPATH="%USERPROFILE%\Applications\Utilities_PC"
attrib -r %TARGETPATH%\*.exe >nul 2>nul
attrib -r %TARGETPATH%\*.dll >nul 2>nul
attrib -r %TARGETPATH%\*.xml >nul 2>nul
attrib -r %TARGETPATH%\*.config >nul 2>nul
del %TARGETPATH%\*.exe >nul 2>nul
del %TARGETPATH%\*.dll >nul 2>nul
del %TARGETPATH%\*.xml >nul 2>nul
del %TARGETPATH%\*.config >nul 2>nul
echo --- Done ---
pause

@echo off
echo Are you sure you want to delete all the User Settings?
echo.
pause
SET TARGETPATH="%USERPROFILE%\Applications\Utilities_PC"
attrib -r %TARGETPATH%\*.settings >nul 2>nul
del %TARGETPATH%\*.settings >nul 2>nul
echo --- Done ---
pause

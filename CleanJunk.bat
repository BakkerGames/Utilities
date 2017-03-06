@echo off

del _delbin.txt >nul 2>nul
dir /ad /s /b . | find "\bin" | find /v "\bin\" | find /v "\." >>_delbin.txt
dir /ad /s /b . | find "\obj" | find /v "\obj\" | find /v "\." >>_delbin.txt
for /f "delims=;" %%a in (_delbin.txt) do rmdir /s /q "%%a"
del _delbin.txt >nul 2>nul

attrib -r -s -h /s *.??proj.user
attrib -r -s -h /s *.bak
attrib -r -s -h /s *.log
attrib -r -s -h /s *.scc
attrib -r -s -h /s *.sln.cache
attrib -r -s -h /s *.suo
attrib -r -s -h /s *.tmp
attrib -r -s -h /s *.vbw
attrib -r -s -h /s *.vspscc
attrib -r -s -h /s *.vssscc
attrib -r -s -h /s build.force

del /s *.??proj.user
del /s *.bak
del /s *.log
del /s *.scc
del /s *.sln.cache
del /s *.suo
del /s *.tmp
del /s *.vbw
del /s *.vspscc
del /s *.vssscc
del /s build.force

pause

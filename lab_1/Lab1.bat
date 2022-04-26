@echo off
cls
:start
echo vvedite rashirenie
set /p r=""
For /R "D:\" %%i In ("%r%") Do (
    (echo %%i >> D:\write.txt)
)
COPY "D:\write.txt" "D:\papka"
start WinRAR A "D:\arhiv.rar" "D:\papka"
pause

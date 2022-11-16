@echo off
for /f "delims=" %%a in ('where /R "%programfiles%\R" /F "Rscript.exe"') do set "p=%%a"
%p% --interactive "path\to\scripts\Destination_Scan_Tjek_script.R"
EXIT

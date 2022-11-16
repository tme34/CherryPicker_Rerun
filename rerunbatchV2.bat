@echo off
for /f "delims=" %%a in ('where /R "%programfiles%\R" /F "Rscript.exe"') do set "p=%%a"
%p% --interactive "path\to\scripts\Rerun_script_V2.R"
EXIT

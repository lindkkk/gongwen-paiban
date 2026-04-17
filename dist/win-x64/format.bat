@echo off
rem =======================================================
rem  gongwen-paiban drag-drop launcher (v3)
rem  Pure ASCII + CRLF. All user-facing Chinese lives in format.ps1.
rem =======================================================

setlocal EnableExtensions

set "LOG=%~dp0paiban-log.txt"
set "PS1=%~dp0format.ps1"
set "EXE=%~dp0gongwen-paiban.exe"

chcp 65001 >nul 2>&1

> "%LOG%" echo ===== gongwen-paiban run %DATE% %TIME% =====
>>"%LOG%" echo BAT path : %~dpnx0
>>"%LOG%" echo PS1 path : %PS1%
>>"%LOG%" echo EXE path : %EXE%
>>"%LOG%" echo Arg1 raw : [%1]
>>"%LOG%" echo Arg1     : [%~1]

echo.
echo ================================================
echo  gongwen-paiban drag-drop launcher
echo  log = %LOG%
echo ================================================
echo.

if "%~1"=="" (
    echo ERROR: No file was dropped onto this bat.
    >>"%LOG%" echo ERROR: no argument
    goto :pause_and_exit
)
if not exist "%EXE%" (
    echo ERROR: gongwen-paiban.exe not found at:
    echo   %EXE%
    >>"%LOG%" echo ERROR: exe missing
    goto :pause_and_exit
)
if not exist "%PS1%" (
    echo ERROR: format.ps1 not found at:
    echo   %PS1%
    >>"%LOG%" echo ERROR: ps1 missing
    goto :pause_and_exit
)
if not exist "%~1" (
    echo ERROR: input file does not exist:
    echo   %~1
    >>"%LOG%" echo ERROR: input file not found
    goto :pause_and_exit
)

echo Running PowerShell interactive launcher...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%PS1%" -InputDocx "%~1"
set "RC=%ERRORLEVEL%"
>>"%LOG%" echo bat saw exit code: %RC%

echo.
if "%RC%"=="0" (
    echo DONE.
) else (
    echo FAILED. Exit code %RC%.
    echo Log:
    echo   %LOG%
)

:pause_and_exit
echo.
echo ================================================
echo  Press any key to close this window...
echo ================================================
pause >nul
exit /b %RC%

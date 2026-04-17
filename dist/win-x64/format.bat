@echo off
rem =======================================================
rem  gongwen-paiban drag-drop launcher
rem  ASCII-only; UTF-8 handling delegated to format.ps1
rem =======================================================

setlocal EnableExtensions

set "LOG=%~dp0paiban-log.txt"
set "PS1=%~dp0format.ps1"
set "EXE=%~dp0gongwen-paiban.exe"

rem Switch console codepage to UTF-8 so any variable expansion
rem containing non-ASCII (e.g. Chinese folder names) is written
rem to the log as UTF-8, matching what PowerShell emits later.
chcp 65001 >nul 2>&1

rem Start the log. Each redirect opens & closes the file, so the
rem log file is NOT held open across the powershell invocation.
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
    echo        Please drag a .docx file onto format.bat.
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
echo   (if no dialog shows up in 10 seconds, see log for details)
echo.

rem No redirection: let ps1 own the log file itself. This avoids
rem Windows exclusive write locks that previously made ps1's own
rem log writes fail silently.
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

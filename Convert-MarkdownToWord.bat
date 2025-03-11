@echo off
setlocal

REM Check if input file is provided
if "%~1"=="" (
    echo Usage: %~nx0 input.txt [output.docx]
    echo Converts a text file with tables to a Word document.
    exit /b 1
)

REM Set input file
set "InputFile=%~1"

REM Determine output file name
if "%~2"=="" (
    set "OutputFile=%~n1.docx"
) else (
    set "OutputFile=%~2"
)

REM Display message
echo Converting %InputFile% to %OutputFile%...

REM Run the PowerShell script
powershell -ExecutionPolicy Bypass -File "%~dp0SimpleTableConverter.ps1" -InputFile "%InputFile%" -OutputFile "%OutputFile%"

REM Check result
if %ERRORLEVEL% EQU 0 (
    echo Conversion complete. Document saved at: %~f2
) else (
    echo Conversion failed with error code %ERRORLEVEL%.
)

exit /b %ERRORLEVEL% 
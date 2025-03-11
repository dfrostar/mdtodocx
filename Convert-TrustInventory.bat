@echo off
setlocal

if "%~1"=="" (
    echo Usage: %~nx0 input.txt [output.docx] [debug]
    echo Specialized converter for trust inventory documents to Word format.
    echo.
    echo Options:
    echo   input.txt  - Path to input trust inventory text file (required)
    echo   output.docx - Path to output Word document (optional)
    echo   debug      - Add this parameter to enable debug output
    exit /b 1
)

set "InputFile=%~1"

REM Determine output file name
if "%~2"=="" (
    set "OutputFile=%~n1-trust.docx"
) else (
    set "OutputFile=%~2"
)

set "DebugOption="

if /i "%~2"=="debug" (
    set "DebugOption=-Debug"
    set "OutputFile=%~n1-trust.docx"
    echo Debug mode enabled.
) else if /i "%~3"=="debug" (
    set "DebugOption=-Debug"
    echo Debug mode enabled.
)

echo Converting Trust Inventory Document: %InputFile% to %OutputFile%
echo Using specialized trust document understanding...

powershell -ExecutionPolicy Bypass -File "%~dp0TrustDocConverter.ps1" -InputFile "%InputFile%" -OutputFile "%OutputFile%" %DebugOption%

if %ERRORLEVEL% EQU 0 (
    echo Conversion complete. Trust document saved at: %OutputFile%
) else (
    echo Conversion failed with error code %ERRORLEVEL%.
)

exit /b %ERRORLEVEL% 
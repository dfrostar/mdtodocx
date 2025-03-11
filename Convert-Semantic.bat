@echo off
setlocal

if "%~1"=="" (
    echo Usage: %~nx0 input.txt output.docx [debug]
    echo Converts a text file to a Word document using semantic understanding.
    echo.
    echo Options:
    echo   input.txt  - Path to input text file (required)
    echo   output.docx - Path to output Word document (required)
    echo   debug      - Add this parameter to enable debug output
    exit /b 1
)

if "%~2"=="" (
    echo Usage: %~nx0 input.txt output.docx [debug]
    echo The output document path is required.
    exit /b 1
)

set "InputFile=%~1"
set "OutputFile=%~2"
set "DebugOption="

if /i "%~3"=="debug" (
    set "DebugOption=-Debug"
    echo Debug mode enabled.
)

echo Converting %InputFile% to %OutputFile% with semantic understanding...

powershell -ExecutionPolicy Bypass -File "%~dp0SemanticDocConverter.ps1" -InputFile "%InputFile%" -OutputFile "%OutputFile%" %DebugOption%

if %ERRORLEVEL% EQU 0 (
    echo Conversion complete. Document saved at: %~f2
) else (
    echo Conversion failed with error code %ERRORLEVEL%.
)

exit /b %ERRORLEVEL% 
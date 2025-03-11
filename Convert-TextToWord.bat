@echo off
setlocal

if "%~1"=="" (
    echo Usage: %~nx0 input.txt output.docx
    echo Converts a text file with tables to a Word document.
    exit /b 1
)

if "%~2"=="" (
    echo Usage: %~nx0 input.txt output.docx
    echo Converts a text file with tables to a Word document.
    exit /b 1
)

powershell -ExecutionPolicy Bypass -File "%~dp0SimpleTableConverter.ps1" -InputFile "%~1" -OutputFile "%~2"

if %ERRORLEVEL% EQU 0 (
    echo Conversion complete. Document saved at: %~f2
) else (
    echo Conversion failed with error code %ERRORLEVEL%.
)

exit /b %ERRORLEVEL% 
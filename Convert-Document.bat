@echo off
setlocal

if "%~1"=="" (
    echo Usage: %~nx0 input_file [output_file] [converter] [show]
    echo Converts a document to Word format using the DocConverter module.
    echo.
    echo Arguments:
    echo   input_file  - Path to input file (required)
    echo   output_file - Path to output Word document (optional)
    echo   converter   - Specific converter to use: Semantic, TrustInventory (optional)
    echo   show        - Add this to show Word during conversion (optional)
    exit /b 1
)

set "InputFile=%~1"
set "OutputFile=%~2"
set "ConverterName=%~3"
set "ShowWord="

if /i "%~3"=="show" (
    set "ShowWord=-ShowWord"
    set "ConverterName="
) else if /i "%~4"=="show" (
    set "ShowWord=-ShowWord"
)

set "VerboseParam="
if defined VERBOSE (
    set "VerboseParam=-Verbose"
)

echo Converting document: %InputFile%
echo Using DocConverter module...

powershell -ExecutionPolicy Bypass -File "%~dp0Convert-Document.ps1" -InputFile "%InputFile%" %OutputFile:!="-OutputFile "%OutputFile%"% %ConverterName:!="-ConverterName %ConverterName%"% %ShowWord% %VerboseParam%

if %ERRORLEVEL% EQU 0 (
    if "%OutputFile%"=="" (
        for %%I in ("%InputFile%") do set "BaseName=%%~nI"
        echo Conversion complete. Document saved at: %~dp1%BaseName%.docx
    ) else (
        echo Conversion complete. Document saved at: %OutputFile%
    )
) else (
    echo Conversion failed with error code %ERRORLEVEL%.
)

exit /b %ERRORLEVEL% 
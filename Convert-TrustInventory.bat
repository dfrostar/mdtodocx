@echo off
setlocal

REM Check if input file is provided
if "%~1"=="" (
    echo Usage: Convert-TrustInventory.bat input_file [output_file] [show]
    echo.
    echo   input_file   - Path to the trust inventory text file
    echo   output_file  - Optional: Path to save the Word document (default: same name with .docx)
    echo   show         - Optional: Add this parameter to show Word during conversion
    exit /b 1
)

set "INPUT_FILE=%~1"

REM Set default output file if not provided
if "%~2"=="" (
    set "OUTPUT_FILE=%~dpn1.docx"
) else (
    set "OUTPUT_FILE=%~2"
)

REM Check if show parameter is provided
set "SHOW_PARAM="
if /i "%~3"=="show" set "SHOW_PARAM=-ShowWord"

echo Converting: %INPUT_FILE% to %OUTPUT_FILE%
if not "%SHOW_PARAM%"=="" echo Word will be visible during conversion

REM Run the PowerShell script with the provided parameters
powershell -ExecutionPolicy Bypass -File "%~dp0Test-TrustInventory.ps1" -InputFile "%INPUT_FILE%" -OutputFile "%OUTPUT_FILE%" %SHOW_PARAM%

if %ERRORLEVEL% neq 0 (
    echo Conversion failed with error code %ERRORLEVEL%
    exit /b %ERRORLEVEL%
)

echo.
echo Conversion completed successfully!
echo Output saved to: %OUTPUT_FILE%
echo.

exit /b 0 
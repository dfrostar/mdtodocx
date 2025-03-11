@echo off
REM Convert-MarkdownToWord.bat - Wrapper for the PowerShell script
REM This batch file makes it easier to run the PowerShell script without having to open PowerShell directly

echo Markdown/TXT to Word Converter
echo ----------------------------
echo.

IF "%~1"=="" (
    echo Usage: Convert-MarkdownToWord.bat input.md [output.docx] [show]
    echo       Convert-MarkdownToWord.bat input.txt [output.docx] [show]
    echo.
    echo Parameters:
    echo   input.md/txt - Path to the input Markdown or TXT file (required)
    echo   output.docx  - Path to the output Word document (optional, defaults to same name as input)
    echo   show         - Add this parameter to show Word during conversion (optional)
    echo.
    echo Examples:
    echo   Convert-MarkdownToWord.bat document.md
    echo   Convert-MarkdownToWord.bat inventory.txt
    echo   Convert-MarkdownToWord.bat document.md result.docx
    echo   Convert-MarkdownToWord.bat document.txt result.docx show
    goto :EOF
)

SET INPUT_FILE=%~1
SET OUTPUT_FILE=%~2
SET SHOW_WORD=%~3

REM If no output file specified, use the input file name with .docx extension
IF "%OUTPUT_FILE%"=="" (
    SET OUTPUT_FILE=%~dpn1.docx
)

REM Prepare the PowerShell command
SET PS_CMD=powershell.exe -ExecutionPolicy Bypass -NoProfile -Command "& '%~dp0Convert-MarkdownToWord.ps1' -InputFile '%INPUT_FILE%' -OutputFile '%OUTPUT_FILE%'"

REM Add ShowWord parameter if specified
IF /I "%SHOW_WORD%"=="show" (
    SET PS_CMD=%PS_CMD% -ShowWord
)

echo Input file:  %INPUT_FILE%
echo Output file: %OUTPUT_FILE%
IF /I "%SHOW_WORD%"=="show" (
    echo Show Word:   Yes
) ELSE (
    echo Show Word:   No
)
echo.
echo Converting...
echo.

REM Execute the PowerShell script
%PS_CMD%

echo.
IF %ERRORLEVEL% EQU 0 (
    echo Conversion completed successfully.
) ELSE (
    echo Conversion failed with error code %ERRORLEVEL%.
)

pause 
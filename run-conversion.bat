@echo off
echo Running text conversion for %1
powershell -ExecutionPolicy Bypass -File "%~dp0Convert-Text.ps1" -InputFile "%1" -OutputFile "%2"
echo Conversion complete! 
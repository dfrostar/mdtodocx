@echo off
echo Converting example files...
echo.

echo 1. Converting the general Markdown document
echo.
Call Convert-Semantic.bat examples\sample-document.md examples\sample-document.docx
echo.

echo 2. Converting the trust inventory document
echo.
Call Convert-TrustInventory.bat examples\sample-trust-inventory.txt examples\sample-trust-inventory.docx
echo.

echo All examples converted successfully!
echo Check the examples directory for the output files.
pause 
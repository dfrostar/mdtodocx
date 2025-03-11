#Requires -Version 5.1
<#
.SYNOPSIS
    Converts text or Markdown documents to properly formatted Word documents.
    
.DESCRIPTION
    This script provides an interface to convert various document formats to Word documents
    with proper formatting, including tables, headings, and structured content.
    It uses semantically-aware converters that analyze the document structure before formatting.
    
.PARAMETER InputFile
    Path to the input file (Markdown, TXT, or other supported formats).
    
.PARAMETER OutputFile
    Path to the output Word document (.docx). If not specified, same name as input with .docx extension.
    
.PARAMETER ConverterName
    Specific converter to use. Options: "Semantic", "TrustInventory", or "Default".
    If not specified, the best converter will be auto-detected.
    
.PARAMETER ShowWord
    Whether to show the Word application during conversion.
    
.PARAMETER Verbose
    Display detailed information about the conversion process.
    
.EXAMPLE
    .\Convert-Document.ps1 -InputFile "document.md" -OutputFile "document.docx"
    
.EXAMPLE
    .\Convert-Document.ps1 -InputFile "trust-inventory.txt" -ConverterName "TrustInventory" -Verbose
    
.LINK
    https://github.com/your-username/DocConverter
#>

param (
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$InputFile,
    
    [Parameter(Mandatory = $false, Position = 1)]
    [string]$OutputFile,
    
    [Parameter(Mandatory = $false)]
    [ValidateSet("Semantic", "TrustInventory", "Default")]
    [string]$ConverterName,
    
    [Parameter(Mandatory = $false)]
    [switch]$ShowWord
)

# Get the directory of this script
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Import the module
$modulePath = Join-Path -Path $scriptDir -ChildPath "DocConverter"
Import-Module $modulePath -Force -Verbose:$VerbosePreference

# Call the converter function
Convert-DocumentToWord -InputFile $InputFile -OutputFile $OutputFile -ConverterName $ConverterName -ShowWord:$ShowWord -Verbose:$VerbosePreference 
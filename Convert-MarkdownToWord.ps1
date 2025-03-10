<#
.SYNOPSIS
    Converts a Markdown file to a properly formatted Microsoft Word document.

.DESCRIPTION
    This PowerShell script takes a Markdown file and converts it to a Microsoft Word document,
    preserving tables, headings, lists, and basic formatting. It uses the Word COM object
    to create and populate the document.

.PARAMETER InputFile
    The path to the Markdown file to convert.

.PARAMETER OutputFile
    The path to save the Word document to.

.PARAMETER ShowWord
    If specified, Word will be visible during the conversion process. Useful for debugging.

.EXAMPLE
    .\Convert-MarkdownToWord.ps1 -InputFile "C:\path\to\markdown.md" -OutputFile "C:\path\to\output.docx"

.EXAMPLE
    .\Convert-MarkdownToWord.ps1 -InputFile "document.md" -OutputFile "document.docx" -ShowWord

.NOTES
    Author: PowerShell Community
    Version: 1.0
    Requires: PowerShell 5.1 or later, Microsoft Word
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$InputFile,
    
    [Parameter(Mandatory = $true)]
    [string]$OutputFile,
    
    [Parameter(Mandatory = $false)]
    [switch]$ShowWord
)

function Process-Markdown {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$InputFile,
        
        [Parameter(Mandatory = $true)]
        [string]$OutputFile,
        
        [Parameter(Mandatory = $false)]
        [switch]$ShowWord
    )
    
    try {
        # Check if the input file exists
        if (-not (Test-Path $InputFile)) {
            Write-Error "The specified input file does not exist: $InputFile"
            return
        }
        
        # Get the absolute path for the output file
        $OutputFile = [System.IO.Path]::GetFullPath($OutputFile)
        
        # Read the markdown content
        $markdownContent = Get-Content -Path $InputFile -Raw
        $markdownLines = Get-Content -Path $InputFile
        
        Write-Host "Converting $InputFile to $OutputFile..."
        
        # Create a new Word application instance
        $word = New-Object -ComObject Word.Application
        $word.Visible = [bool]$ShowWord
        
        # Create a new document
        $doc = $word.Documents.Add()
        $selection = $word.Selection
        
        # Track whether we're currently in a table
        $inTable = $false
        $tableRows = @()
        $tableHeaders = @()
        
        # Track if we're in a list
        $inList = $false
        $listType = $null # "ordered" or "unordered"
        
        # Process each line
        for ($i = 0; $i -lt $markdownLines.Count; $i++) {
            $line = $markdownLines[$i]
            
            # Skip empty lines unless they indicate the end of a table
            if ([string]::IsNullOrWhiteSpace($line)) {
                if ($inTable) {
                    # End of table
                    Add-Table -Selection $selection -Headers $tableHeaders -Rows $tableRows
                    $inTable = $false
                    $tableRows = @()
                    $tableHeaders = @()
                }
                elseif ($inList) {
                    # End of list
                    $inList = $false
                    $listType = $null
                    $selection.TypeParagraph()
                }
                else {
                    # Regular paragraph break
                    $selection.TypeParagraph()
                }
                continue
            }
            
            # Process headings (# Heading)
            if ($line -match '^(#{1,6})\s+(.+)$') {
                $headingLevel = $matches[1].Length
                $headingText = $matches[2]
                
                # End table if in progress
                if ($inTable) {
                    Add-Table -Selection $selection -Headers $tableHeaders -Rows $tableRows
                    $inTable = $false
                    $tableRows = @()
                    $tableHeaders = @()
                }
                
                # Apply heading style
                $selection.Style = $word.ActiveDocument.Styles.Item("Heading $headingLevel")
                $selection.TypeText($headingText)
                $selection.TypeParagraph()
                continue
            }
            
            # Process table header row (| Header1 | Header2 |)
            if ($line -match '^\s*\|(.+)\|\s*$') {
                $cells = $matches[1] -split '\|' | ForEach-Object { $_.Trim() }
                
                # If this is the first row, it's the header
                if (-not $inTable) {
                    $inTable = $true
                    $tableHeaders = $cells
                }
                # Otherwise, it's a data row (if it's not the separator row)
                elseif (-not ($line -match '^\s*\|(\s*[-:]+[-|\s:]*)+\|\s*$')) {
                    $tableRows += , $cells
                }
                continue
            }
            
            # Process unordered list items (- Item or * Item)
            if ($line -match '^\s*[-*]\s+(.+)$') {
                $listItem = $matches[1]
                
                if (-not $inList -or $listType -ne "unordered") {
                    # Start a new unordered list
                    $inList = $true
                    $listType = "unordered"
                    $selection.Range.ListFormat.ApplyBulletDefault()
                }
                
                # Add the list item
                $selection.TypeText($listItem)
                $selection.TypeParagraph()
                continue
            }
            
            # Process ordered list items (1. Item)
            if ($line -match '^\s*\d+\.\s+(.+)$') {
                $listItem = $matches[1]
                
                if (-not $inList -or $listType -ne "ordered") {
                    # Start a new ordered list
                    $inList = $true
                    $listType = "ordered"
                    $selection.Range.ListFormat.ApplyNumberDefault()
                }
                
                # Add the list item
                $selection.TypeText($listItem)
                $selection.TypeParagraph()
                continue
            }
            
            # If we get here, it's regular text
            if ($inList) {
                # End the list
                $selection.Range.ListFormat.RemoveNumbers()
                $inList = $false
                $listType = $null
            }
            
            # Process bold and italic formatting
            $formattedLine = $line
            
            # First, handle Bold+Italic (**_text_**)
            $boldItalicMatches = [regex]::Matches($formattedLine, '\*\*\*(.+?)\*\*\*|\*\*_(.+?)_\*\*|__\*(.+?)\*__|_\*\*(.+?)\*\*_')
            foreach ($match in $boldItalicMatches) {
                # Find the actual match group that captured text
                $capturedText = $null
                for ($j = 1; $j -lt $match.Groups.Count; $j++) {
                    if ($match.Groups[$j].Success) {
                        $capturedText = $match.Groups[$j].Value
                        break
                    }
                }
                
                if ($capturedText) {
                    $selection.TypeText($formattedLine.Substring(0, $match.Index))
                    $selection.Font.Bold = 1
                    $selection.Font.Italic = 1
                    $selection.TypeText($capturedText)
                    $selection.Font.Bold = 0
                    $selection.Font.Italic = 0
                    $formattedLine = $formattedLine.Substring($match.Index + $match.Length)
                }
            }
            
            # Handle Bold (**text**)
            $boldMatches = [regex]::Matches($formattedLine, '\*\*(.+?)\*\*|__(.+?)__')
            foreach ($match in $boldMatches) {
                # Find the actual match group that captured text
                $capturedText = $null
                for ($j = 1; $j -lt $match.Groups.Count; $j++) {
                    if ($match.Groups[$j].Success) {
                        $capturedText = $match.Groups[$j].Value
                        break
                    }
                }
                
                if ($capturedText) {
                    $selection.TypeText($formattedLine.Substring(0, $match.Index))
                    $selection.Font.Bold = 1
                    $selection.TypeText($capturedText)
                    $selection.Font.Bold = 0
                    $formattedLine = $formattedLine.Substring($match.Index + $match.Length)
                }
            }
            
            # Handle Italic (*text*)
            $italicMatches = [regex]::Matches($formattedLine, '\*(.+?)\*|_(.+?)_')
            foreach ($match in $italicMatches) {
                # Find the actual match group that captured text
                $capturedText = $null
                for ($j = 1; $j -lt $match.Groups.Count; $j++) {
                    if ($match.Groups[$j].Success) {
                        $capturedText = $match.Groups[$j].Value
                        break
                    }
                }
                
                if ($capturedText) {
                    $selection.TypeText($formattedLine.Substring(0, $match.Index))
                    $selection.Font.Italic = 1
                    $selection.TypeText($capturedText)
                    $selection.Font.Italic = 0
                    $formattedLine = $formattedLine.Substring($match.Index + $match.Length)
                }
            }
            
            # Type any remaining text
            if ($formattedLine.Length -gt 0) {
                $selection.TypeText($formattedLine)
            }
            
            $selection.TypeParagraph()
        }
        
        # Check if we need to finalize a table
        if ($inTable) {
            Add-Table -Selection $selection -Headers $tableHeaders -Rows $tableRows
        }
        
        # Save the document
        $doc.SaveAs([ref]$OutputFile, [ref]16) # 16 = wdFormatDocumentDefault (*.docx)
        $doc.Close()
        
        # Close Word if it was not visible
        if (-not $ShowWord) {
            $word.Quit()
        }
        
        Write-Host "Conversion complete. Document saved to: $OutputFile" -ForegroundColor Green
    }
    catch {
        Write-Error "An error occurred during the conversion: $_"
        
        # Attempt to clean up
        if ($doc -ne $null) {
            try { $doc.Close([ref]$false) } catch { }
        }
        
        if ($word -ne $null -and -not $ShowWord) {
            try { $word.Quit() } catch { }
        }
    }
    finally {
        # Release COM objects
        if ($selection -ne $null) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($selection) | Out-Null }
        if ($doc -ne $null) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null }
        if ($word -ne $null) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

function Add-Table {
    [CmdletBinding()]
    param (
        $Selection,
        [string[]]$Headers,
        [array]$Rows
    )
    
    # Determine the number of columns and rows
    $columnCount = $Headers.Count
    $rowCount = $Rows.Count
    
    # Create a table with headers + data rows
    $table = $Selection.Tables.Add($Selection.Range, $rowCount + 1, $columnCount)
    
    # Set table properties for better appearance
    $table.Borders.Enable = $true
    $table.Style = "Grid Table 4 - Accent 1"
    
    # Add headers
    for ($col = 0; $col -lt $columnCount; $col++) {
        $table.Cell(1, $col + 1).Range.Text = $Headers[$col]
    }
    
    # Make the header row bold
    $table.Rows.Item(1).Range.Font.Bold = $true
    
    # Add data rows
    for ($row = 0; $row -lt $rowCount; $row++) {
        $rowData = $Rows[$row]
        for ($col = 0; $col -lt [Math]::Min($columnCount, $rowData.Count); $col++) {
            $cellText = $rowData[$col]
            if (-not [string]::IsNullOrEmpty($cellText)) {
                $table.Cell($row + 2, $col + 1).Range.Text = $cellText
            }
        }
    }
    
    # Add a paragraph after the table
    $Selection.TypeParagraph()
}

# Execute the script
Process-Markdown -InputFile $InputFile -OutputFile $OutputFile -ShowWord:$ShowWord 
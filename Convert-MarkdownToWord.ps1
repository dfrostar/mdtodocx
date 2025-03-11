<#
.SYNOPSIS
    Converts a Markdown or TXT file to a properly formatted Microsoft Word document.

.DESCRIPTION
    This PowerShell script takes a Markdown or TXT file and converts it to a Microsoft Word document,
    preserving tables, headings, lists, and basic formatting. It uses the Word COM object
    to create and populate the document.

.PARAMETER InputFile
    The path to the Markdown or TXT file to convert.

.PARAMETER OutputFile
    The path to save the Word document to.

.PARAMETER ShowWord
    If specified, Word will be visible during the conversion process. Useful for debugging.

.EXAMPLE
    .\Convert-MarkdownToWord.ps1 -InputFile "C:\path\to\markdown.md" -OutputFile "C:\path\to\output.docx"

.EXAMPLE
    .\Convert-MarkdownToWord.ps1 -InputFile "C:\path\to\data.txt" -OutputFile "C:\path\to\output.docx"

.EXAMPLE
    .\Convert-MarkdownToWord.ps1 -InputFile "document.md" -OutputFile "document.docx" -ShowWord

.NOTES
    Author: PowerShell Community
    Version: 1.1
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

function Process-Document {
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
        
        # Get the file extension to determine if it's markdown or text
        $fileExtension = [System.IO.Path]::GetExtension($InputFile).ToLower()
        $isMarkdown = $fileExtension -eq ".md"
        $isTxt = $fileExtension -eq ".txt"
        
        if (-not ($isMarkdown -or $isTxt)) {
            Write-Error "The input file must be a Markdown (.md) or Text (.txt) file. File extension: $fileExtension"
            return
        }
        
        # Log file for debugging
        $logPath = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($OutputFile), "conversion_log.txt")
        "Starting conversion of $InputFile to $OutputFile at $(Get-Date)" | Out-File -FilePath $logPath
        
        $fileType = if ($isMarkdown) { "Markdown" } else { "Text" }
        
        # Get the absolute path for the output file
        $OutputFile = [System.IO.Path]::GetFullPath($OutputFile)
        
        # Read the document content
        $documentContent = Get-Content -Path $InputFile -Raw
        $documentLines = Get-Content -Path $InputFile
        
        "Read $($documentLines.Count) lines from $InputFile" | Out-File -FilePath $logPath -Append
        
        Write-Host "Converting $fileType file $InputFile to $OutputFile..."
        
        # Create a new Word application instance
        "Creating Word application instance" | Out-File -FilePath $logPath -Append
        $word = New-Object -ComObject Word.Application
        $word.Visible = [bool]$ShowWord
        
        # Create a new document
        $doc = $word.Documents.Add()
        $selection = $word.Selection
        
        "Processing document content" | Out-File -FilePath $logPath -Append
        
        # For TXT files, we'll use a completely different table detection approach
        if ($isTxt) {
            "Using TXT-specific processing method" | Out-File -FilePath $logPath -Append
            Process-TxtDocument -DocumentLines $documentLines -Selection $selection -LogPath $logPath
        } else {
            "Using Markdown processing method" | Out-File -FilePath $logPath -Append
            Process-MarkdownDocument -DocumentLines $documentLines -Selection $selection -LogPath $logPath
        }
        
        # Save the document
        "Saving document to $OutputFile" | Out-File -FilePath $logPath -Append
        $doc.SaveAs([ref]$OutputFile, [ref]16) # 16 = wdFormatDocumentDefault (*.docx)
        $doc.Close()
        
        # Close Word if it was not visible
        if (-not $ShowWord) {
            $word.Quit()
        }
        
        Write-Host "Conversion complete. Document saved to: $OutputFile" -ForegroundColor Green
        "Conversion completed successfully at $(Get-Date)" | Out-File -FilePath $logPath -Append
    }
    catch {
        $errorMessage = "An error occurred during the conversion: $_"
        Write-Error $errorMessage
        $errorMessage | Out-File -FilePath $logPath -Append
        
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

function Process-TxtDocument {
    param (
        [string[]]$DocumentLines,
        $Selection,
        [string]$LogPath
    )
    
    # Track whether we're currently in a table
    $inTable = $false
    $tableSectionLines = @()
    $currentHeading = $null
    
    # Process each line
    for ($i = 0; $i -lt $DocumentLines.Count; $i++) {
        $line = $DocumentLines[$i]
        
        # Log processing information
        "Processing line $i`: '$line'" | Out-File -FilePath $LogPath -Append
        
        # Skip empty lines unless they indicate the end of a table
        if ([string]::IsNullOrWhiteSpace($line)) {
            if ($inTable -and $tableSectionLines.Count -gt 0) {
                # End of table, process collected table lines
                "Table section ended, processing collected lines with $($tableSectionLines.Count) lines" | Out-File -FilePath $LogPath -Append
                Process-TableSection -TableLines $tableSectionLines -Selection $Selection -LogPath $LogPath
                $inTable = $false
                $tableSectionLines = @()
            }
            
            $Selection.TypeParagraph()
            continue
        }
        
        # Process headings (# Heading or ## Heading)
        if ($line -match '^(#{1,6})\s+(.+)$') {
            $headingLevel = $matches[1].Length
            $headingText = $matches[2]
            
            # End table if in progress
            if ($inTable -and $tableSectionLines.Count -gt 0) {
                "Table section ended due to heading, processing collected lines" | Out-File -FilePath $LogPath -Append
                Process-TableSection -TableLines $tableSectionLines -Selection $Selection -LogPath $LogPath
                $inTable = $false
                $tableSectionLines = @()
            }
            
            $currentHeading = $headingText
            "Found heading level $headingLevel`: $headingText" | Out-File -FilePath $LogPath -Append
            
            # Apply heading style
            $Selection.Style = $Selection.Document.Styles.Item("Heading $headingLevel")
            $Selection.TypeText($headingText)
            $Selection.TypeParagraph()
            continue
        }
        
        # Check if this line starts with a pipe (potential table row)
        if ($line -match '^\s*\|(.+)\|\s*$') {
            if (-not $inTable) {
                # Start of a new pipe-delimited section
                $inTable = $true
                "Found start of pipe-delimited section" | Out-File -FilePath $LogPath -Append
            }
            
            $tableSectionLines += $line
            continue
        }
        # If we're already in a table and this line doesn't start with a pipe, check if it could be a continuation
        elseif ($inTable -and -not ($line -match '^\s*\|')) {
            # If it's a separator line made of dashes or plus signs, keep it in the table
            if ($line -match '^\s*[-+]+\s*$') {
                $tableSectionLines += $line
                continue
            }
            
            # Otherwise, it's the end of the table section
            "Table section ended due to non-pipe line, processing collected lines" | Out-File -FilePath $LogPath -Append
            Process-TableSection -TableLines $tableSectionLines -Selection $Selection -LogPath $LogPath
            $inTable = $false
            $tableSectionLines = @()
            
            # Process this line as regular text
            "Processing as regular text: $line" | Out-File -FilePath $LogPath -Append
            $Selection.TypeText($line)
            $Selection.TypeParagraph()
            continue
        }
        
        # If we're here, it's regular text that's not part of a table
        if ($line -match '^\s*[-*]\s+(.+)$') {
            # This is a bullet point
            $listItem = $matches[1]
            "Found bullet point: $listItem" | Out-File -FilePath $LogPath -Append
            
            $Selection.Range.ListFormat.ApplyBulletDefault()
            $Selection.TypeText($listItem)
            $Selection.TypeParagraph()
        }
        elseif ($line -match '^\s*\d+\.\s+(.+)$') {
            # This is a numbered list item
            $listItem = $matches[1]
            "Found numbered list item: $listItem" | Out-File -FilePath $LogPath -Append
            
            $Selection.Range.ListFormat.ApplyNumberDefault()
            $Selection.TypeText($listItem)
            $Selection.TypeParagraph()
        }
        else {
            # Regular text
            "Processing as regular text: $line" | Out-File -FilePath $LogPath -Append
            $Selection.TypeText($line)
            $Selection.TypeParagraph()
        }
    }
    
    # Check if we need to finalize a table section
    if ($inTable -and $tableSectionLines.Count -gt 0) {
        "Finalizing last table section with $($tableSectionLines.Count) lines" | Out-File -FilePath $LogPath -Append
        Process-TableSection -TableLines $tableSectionLines -Selection $Selection -LogPath $LogPath
    }
}

function Process-TableSection {
    param (
        [string[]]$TableLines,
        $Selection,
        [string]$LogPath
    )
    
    "Processing table section with $($TableLines.Count) lines" | Out-File -FilePath $LogPath -Append
    
    # Find all pipe-delimited lines
    $pipeDelimitedLines = @()
    foreach ($line in $TableLines) {
        if ($line -match '^\s*\|(.+)\|\s*$') {
            $pipeDelimitedLines += $line
            "Found pipe-delimited line: $line" | Out-File -FilePath $LogPath -Append
        }
    }
    
    # If we found less than 1 pipe-delimited line, just output as text
    if ($pipeDelimitedLines.Count -lt 1) {
        "Not enough pipe-delimited lines to form a table, treating as text" | Out-File -FilePath $LogPath -Append
        foreach ($line in $TableLines) {
            $Selection.TypeText($line)
            $Selection.TypeParagraph()
        }
        return
    }
    
    # The first pipe-delimited line is the header
    $headerRow = $pipeDelimitedLines[0]
    
    # Check if there's a separator row (all dashes and pipes)
    $separatorRowIndex = -1
    for ($i = 1; $i -lt $pipeDelimitedLines.Count; $i++) {
        if ($pipeDelimitedLines[$i] -match '^\s*\|([-|]+)\|\s*$') {
            $separatorRowIndex = $i
            "Found separator row at index $i" | Out-File -FilePath $LogPath -Append
            break
        }
    }
    
    # Data rows start after the separator row, or from index 1 if no separator row
    $dataStartIndex = if ($separatorRowIndex -gt 0) { $separatorRowIndex + 1 } else { 1 }
    $dataRows = $pipeDelimitedLines[$dataStartIndex..$pipeDelimitedLines.Count]
    
    "Header row: $headerRow" | Out-File -FilePath $LogPath -Append
    "Found $($dataRows.Count) data rows" | Out-File -FilePath $LogPath -Append
    
    # Extract header cells by careful parsing the pipe-delimited line
    $headerCells = @()
    
    # First, clean up the header row by removing outer pipes
    $trimmedHeader = $headerRow.Trim()
    if ($trimmedHeader.StartsWith("|")) {
        $trimmedHeader = $trimmedHeader.Substring(1)
    }
    if ($trimmedHeader.EndsWith("|")) {
        $trimmedHeader = $trimmedHeader.Substring(0, $trimmedHeader.Length - 1)
    }
    
    # Split by pipe, being sure to preserve empty cells, and trim each cell
    $headerCells = $trimmedHeader -split '\|' | ForEach-Object { $_.Trim() }
    
    "Found $($headerCells.Count) header cells:" | Out-File -FilePath $LogPath -Append
    foreach ($cell in $headerCells) {
        "  Header cell: '$cell'" | Out-File -FilePath $LogPath -Append
    }
    
    # Extract data rows
    $tableDataRows = @()
    foreach ($dataRow in $dataRows) {
        # Skip if it's a separator row (all dashes and pipes)
        if ($dataRow -match '^\s*\|([-|]+)\|\s*$') {
            "Skipping separator row: $dataRow" | Out-File -FilePath $LogPath -Append
            continue
        }
        
        # Clean up and split the row
        $trimmedRow = $dataRow.Trim()
        if ($trimmedRow.StartsWith("|")) {
            $trimmedRow = $trimmedRow.Substring(1)
        }
        if ($trimmedRow.EndsWith("|")) {
            $trimmedRow = $trimmedRow.Substring(0, $trimmedRow.Length - 1)
        }
        
        # Split by pipe and ensure we maintain the same number of cells as headers
        $rowCells = $trimmedRow -split '\|'
        
        # Trim whitespace from each cell
        $rowCells = $rowCells | ForEach-Object { $_.Trim() }
        
        # If we need more cells to match header count, add empty strings
        while ($rowCells.Count -lt $headerCells.Count) {
            $rowCells += ""
        }
        
        # If we have too many cells, trim to match header count
        if ($rowCells.Count -gt $headerCells.Count) {
            $rowCells = $rowCells[0..($headerCells.Count - 1)]
        }
        
        $tableDataRows += , $rowCells
        
        "Added data row with $($rowCells.Count) cells" | Out-File -FilePath $LogPath -Append
        foreach ($cell in $rowCells) {
            "  Data cell: '$cell'" | Out-File -FilePath $LogPath -Append
        }
    }
    
    # Create the table in Word
    "Creating Word table with $($headerCells.Count) columns and $($tableDataRows.Count) data rows" | Out-File -FilePath $LogPath -Append
    Create-SimpleTable -HeaderCells $headerCells -DataRows $tableDataRows -Selection $Selection -LogPath $LogPath
}

function Create-SimpleTable {
    param (
        [string[]]$HeaderCells,
        [array]$DataRows,
        $Selection,
        [string]$LogPath
    )
    
    # Ensure we have headers
    if ($HeaderCells.Count -eq 0) {
        "No header cells found, cannot create table" | Out-File -FilePath $LogPath -Append
        return
    }
    
    $columnCount = $HeaderCells.Count
    $rowCount = $DataRows.Count
    
    "Creating simple table with $columnCount columns and $rowCount rows" | Out-File -FilePath $LogPath -Append
    
    try {
        # Create a table
        $table = $Selection.Tables.Add($Selection.Range, $rowCount + 1, $columnCount)
        
        # Apply basic formatting
        $table.Borders.Enable = $true
        $table.AllowAutoFit = $true
        
        # Set table style to a simple grid
        $table.Style = "Table Grid"
        
        # Clear all cell shading
        for ($row = 1; $row -le $rowCount + 1; $row++) {
            for ($col = 1; $col -le $columnCount; $col++) {
                try {
                    $table.Cell($row, $col).Shading.BackgroundPatternColor = -16777216  # wdColorAutomatic
                } catch {
                    "Error setting cell shading: $_" | Out-File -FilePath $LogPath -Append
                }
            }
        }
        
        # Add header cells
        for ($col = 0; $col -lt $columnCount; $col++) {
            if ($col -lt $HeaderCells.Count) {
                try {
                    $cellText = $HeaderCells[$col]
                    $table.Cell(1, $col + 1).Range.Text = $cellText
                    "Added header cell (1, $($col+1)): '$cellText'" | Out-File -FilePath $LogPath -Append
                } catch {
                    "Error adding header cell: $_" | Out-File -FilePath $LogPath -Append
                }
            }
        }
        
        # Make header row bold
        $table.Rows.Item(1).Range.Bold = $true
        
        # Add data rows
        for ($row = 0; $row -lt $rowCount; $row++) {
            $rowData = $DataRows[$row]
            
            # Calculate how many cells to add
            $cellsToAdd = [Math]::Min($columnCount, $rowData.Count)
            
            for ($col = 0; $col -lt $cellsToAdd; $col++) {
                try {
                    $cellText = $rowData[$col]
                    if (-not [string]::IsNullOrEmpty($cellText)) {
                        $table.Cell($row + 2, $col + 1).Range.Text = $cellText
                        "Added data cell ($($row+2), $($col+1)): '$cellText'" | Out-File -FilePath $LogPath -Append
                    }
                } catch {
                    "Error adding data cell: $_" | Out-File -FilePath $LogPath -Append
                }
            }
        }
        
        # Set column widths to fit content
        try {
            $table.AutoFitBehavior(1)  # wdAutoFitContent = 1
            "Applied AutoFit to table" | Out-File -FilePath $LogPath -Append
        } catch {
            "Error applying AutoFit: $_" | Out-File -FilePath $LogPath -Append
        }
        
    } catch {
        "Error creating table: $_" | Out-File -FilePath $LogPath -Append
    }
    
    # Add a paragraph after the table
    $Selection.TypeParagraph()
}

function Process-MarkdownDocument {
    param (
        [string[]]$DocumentLines,
        $Selection,
        [string]$LogPath
    )
    
    # Track whether we're currently in a table
    $inTable = $false
    $tableRows = @()
    $tableHeaders = @()
    
    # Track if we're in a list
    $inList = $false
    $listType = $null # "ordered" or "unordered"
    
    # Process each line
    for ($i = 0; $i -lt $DocumentLines.Count; $i++) {
        $line = $DocumentLines[$i]
        
        # Skip empty lines unless they indicate the end of a table
        if ([string]::IsNullOrWhiteSpace($line)) {
            if ($inTable) {
                # End of table
                Add-Table -Selection $Selection -Headers $tableHeaders -Rows $tableRows
                $inTable = $false
                $tableRows = @()
                $tableHeaders = @()
            }
            elseif ($inList) {
                # End of list
                $inList = $false
                $listType = $null
                $Selection.TypeParagraph()
            }
            else {
                # Regular paragraph break
                $Selection.TypeParagraph()
            }
            continue
        }
        
        # Process headings (# Heading)
        if ($line -match '^(#{1,6})\s+(.+)$') {
            $headingLevel = $matches[1].Length
            $headingText = $matches[2]
            
            # End table if in progress
            if ($inTable) {
                Add-Table -Selection $Selection -Headers $tableHeaders -Rows $tableRows
                $inTable = $false
                $tableRows = @()
                $tableHeaders = @()
            }
            
            # Apply heading style
            $Selection.Style = $Selection.Document.Styles.Item("Heading $headingLevel")
            $Selection.TypeText($headingText)
            $Selection.TypeParagraph()
            continue
        }
        
        # Process table row (| Column1 | Column2 |)
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
                $Selection.Range.ListFormat.ApplyBulletDefault()
            }
            
            # Add the list item
            $Selection.TypeText($listItem)
            $Selection.TypeParagraph()
            continue
        }
        
        # Process ordered list items (1. Item)
        if ($line -match '^\s*\d+\.\s+(.+)$') {
            $listItem = $matches[1]
            
            if (-not $inList -or $listType -ne "ordered") {
                # Start a new ordered list
                $inList = $true
                $listType = "ordered"
                $Selection.Range.ListFormat.ApplyNumberDefault()
            }
            
            # Add the list item
            $Selection.TypeText($listItem)
            $Selection.TypeParagraph()
            continue
        }
        
        # If we get here, it's regular text
        if ($inList) {
            # End the list
            $Selection.Range.ListFormat.RemoveNumbers()
            $inList = $false
            $listType = $null
        }
        
        # Process formatting and type text
        Process-FormattedText -Line $line -Selection $Selection
    }
    
    # Check if we need to finalize a table
    if ($inTable) {
        Add-Table -Selection $Selection -Headers $tableHeaders -Rows $tableRows
    }
}

function Process-FormattedText {
    param (
        [string]$Line,
        $Selection
    )
    
    # Handle formatting for bold, italic, etc.
    $formattedLine = $Line
    
    # First, handle Bold+Italic (**_text_**)
    $boldItalicMatches = [regex]::Matches($formattedLine, '\*\*\*(.+?)\*\*\*|\*\*_(.+?)_\*\*|__\*(.+?)\*__|_\*\*(.+?)\*\*_')
    foreach ($match in $boldItalicMatches) {
        $capturedText = $null
        for ($j = 1; $j -lt $match.Groups.Count; $j++) {
            if ($match.Groups[$j].Success) {
                $capturedText = $match.Groups[$j].Value
                break
            }
        }
        
        if ($capturedText) {
            $Selection.TypeText($formattedLine.Substring(0, $match.Index))
            $Selection.Font.Bold = 1
            $Selection.Font.Italic = 1
            $Selection.TypeText($capturedText)
            $Selection.Font.Bold = 0
            $Selection.Font.Italic = 0
            $formattedLine = $formattedLine.Substring($match.Index + $match.Length)
        }
    }
    
    # Handle Bold (**text**)
    $boldMatches = [regex]::Matches($formattedLine, '\*\*(.+?)\*\*|__(.+?)__')
    foreach ($match in $boldMatches) {
        $capturedText = $null
        for ($j = 1; $j -lt $match.Groups.Count; $j++) {
            if ($match.Groups[$j].Success) {
                $capturedText = $match.Groups[$j].Value
                break
            }
        }
        
        if ($capturedText) {
            $Selection.TypeText($formattedLine.Substring(0, $match.Index))
            $Selection.Font.Bold = 1
            $Selection.TypeText($capturedText)
            $Selection.Font.Bold = 0
            $formattedLine = $formattedLine.Substring($match.Index + $match.Length)
        }
    }
    
    # Handle Italic (*text*)
    $italicMatches = [regex]::Matches($formattedLine, '\*(.+?)\*|_(.+?)_')
    foreach ($match in $italicMatches) {
        $capturedText = $null
        for ($j = 1; $j -lt $match.Groups.Count; $j++) {
            if ($match.Groups[$j].Success) {
                $capturedText = $match.Groups[$j].Value
                break
            }
        }
        
        if ($capturedText) {
            $Selection.TypeText($formattedLine.Substring(0, $match.Index))
            $Selection.Font.Italic = 1
            $Selection.TypeText($capturedText)
            $Selection.Font.Italic = 0
            $formattedLine = $formattedLine.Substring($match.Index + $match.Length)
        }
    }
    
    # Type any remaining text
    if ($formattedLine.Length -gt 0) {
        $Selection.TypeText($formattedLine)
    }
    
    $Selection.TypeParagraph()
}

function Create-TableFromText {
    param (
        [string[]]$TableLines,
        $Selection,
        [string]$LogPath
    )
    
    "Starting to create table from $($TableLines.Count) lines" | Out-File -FilePath $LogPath -Append
    
    # Identify header row and data rows
    $headerRowIndex = -1
    $headerSeparatorIndex = -1
    $dataRowIndices = @()
    
    for ($i = 0; $i -lt $TableLines.Count; $i++) {
        $line = $TableLines[$i]
        
        if ($line -match '^\s*\+[-+]+\+\s*$') {
            # This is a separator line with plus signs
            if ($headerRowIndex -eq -1) {
                # This is the top border of the table
                continue
            } 
            elseif ($headerSeparatorIndex -eq -1) {
                # This is the separator after the header
                $headerSeparatorIndex = $i
            }
        }
        elseif ($line -match '^\s*\|(.+)\|\s*$') {
            # This is a content line
            if ($headerRowIndex -eq -1) {
                $headerRowIndex = $i
                "Found header row at index $i" | Out-File -FilePath $LogPath -Append
            } else {
                $dataRowIndices += $i
                "Found data row at index $i" | Out-File -FilePath $LogPath -Append
            }
        }
    }
    
    # If we didn't find a proper header and data structure, treat as text
    if ($headerRowIndex -eq -1 -or $dataRowIndices.Count -eq 0) {
        "Could not identify table structure, treating as text" | Out-File -FilePath $LogPath -Append
        foreach ($line in $TableLines) {
            $Selection.TypeText($line)
            $Selection.TypeParagraph()
        }
        return
    }
    
    # Extract header cells from the header row
    $headerRow = $TableLines[$headerRowIndex]
    $headerCells = @()
    
    # Extract cells by manually parsing the pipe-delimited line
    # First, remove the outer pipes
    $trimmedHeader = $headerRow.Trim()
    if ($trimmedHeader.StartsWith("|")) {
        $trimmedHeader = $trimmedHeader.Substring(1)
    }
    if ($trimmedHeader.EndsWith("|")) {
        $trimmedHeader = $trimmedHeader.Substring(0, $trimmedHeader.Length - 1)
    }
    
    # Split by pipe and trim each cell
    $headerCells = $trimmedHeader -split '\|' | ForEach-Object { $_.Trim() }
    
    "Found $($headerCells.Count) header cells" | Out-File -FilePath $LogPath -Append
    foreach ($cell in $headerCells) {
        "  Header cell: '$cell'" | Out-File -FilePath $LogPath -Append
    }
    
    # Extract data rows
    $dataRows = @()
    foreach ($rowIndex in $dataRowIndices) {
        $dataRow = $TableLines[$rowIndex]
        
        # Extract cells by manually parsing the pipe-delimited line
        # First, remove the outer pipes
        $trimmedRow = $dataRow.Trim()
        if ($trimmedRow.StartsWith("|")) {
            $trimmedRow = $trimmedRow.Substring(1)
        }
        if ($trimmedRow.EndsWith("|")) {
            $trimmedRow = $trimmedRow.Substring(0, $trimmedRow.Length - 1)
        }
        
        # Split by pipe and trim each cell
        $rowCells = $trimmedRow -split '\|' | ForEach-Object { $_.Trim() }
        $dataRows += , $rowCells
        
        "Found data row with $($rowCells.Count) cells" | Out-File -FilePath $LogPath -Append
        foreach ($cell in $rowCells) {
            "  Data cell: '$cell'" | Out-File -FilePath $LogPath -Append
        }
    }
    
    # Create the table in Word
    "Creating Word table with $($headerCells.Count) columns and $($dataRows.Count) data rows" | Out-File -FilePath $LogPath -Append
    Create-WordTable -HeaderCells $headerCells -DataRows $dataRows -Selection $Selection -LogPath $LogPath
}

function Create-WordTable {
    param (
        [string[]]$HeaderCells,
        [array]$DataRows,
        $Selection,
        [string]$LogPath
    )
    
    # Ensure we have data
    if ($HeaderCells.Count -eq 0) {
        "No header cells found, cannot create table" | Out-File -FilePath $LogPath -Append
        return
    }
    
    $columnCount = $HeaderCells.Count
    $rowCount = $DataRows.Count
    
    "Creating table with $columnCount columns and $rowCount rows" | Out-File -FilePath $LogPath -Append
    
    # Create a table with headers + data rows
    $table = $Selection.Tables.Add($Selection.Range, $rowCount + 1, $columnCount)
    
    # Set table properties for better appearance
    $table.Borders.Enable = $true
    
    # Set grid style without shading
    try {
        "Setting table style to Basic Grid" | Out-File -FilePath $LogPath -Append
        $table.Style = "Table Grid"
        
        "Removing all shading from table" | Out-File -FilePath $LogPath -Append
        for ($row = 1; $row -le $rowCount + 1; $row++) {
            for ($col = 1; $col -le $columnCount; $col++) {
                # Set cell shading to none (wdColorAutomatic = -16777216)
                $table.Cell($row, $col).Shading.BackgroundPatternColor = -16777216
            }
        }
    }
    catch {
        "Error setting table style: $_" | Out-File -FilePath $LogPath -Append
    }
    
    # Add headers
    "Adding header cells to table" | Out-File -FilePath $LogPath -Append
    for ($col = 0; $col -lt $columnCount; $col++) {
        if ($col -lt $HeaderCells.Count) {
            $cellText = $HeaderCells[$col]
            "Setting header cell ($1, $($col+1)) to: '$cellText'" | Out-File -FilePath $LogPath -Append
            $table.Cell(1, $col + 1).Range.Text = $cellText
        }
    }
    
    # Make the header row bold
    "Making header row bold" | Out-File -FilePath $LogPath -Append
    $table.Rows.Item(1).Range.Bold = $true
    
    # Add data rows
    "Adding data rows to table" | Out-File -FilePath $LogPath -Append
    for ($row = 0; $row -lt $rowCount; $row++) {
        $rowData = $DataRows[$row]
        for ($col = 0; $col -lt [Math]::Min($columnCount, $rowData.Count); $col++) {
            $cellText = $rowData[$col]
            if (-not [string]::IsNullOrEmpty($cellText)) {
                "Setting data cell ($($row+2), $($col+1)) to: '$cellText'" | Out-File -FilePath $LogPath -Append
                $table.Cell($row + 2, $col + 1).Range.Text = $cellText
            }
        }
    }
    
    # Auto-fit to contents
    "Applying AutoFit to table contents" | Out-File -FilePath $LogPath -Append
    try {
        $table.AutoFitBehavior(1) # wdAutoFitContent = 1
    }
    catch {
        "Error applying AutoFit: $_" | Out-File -FilePath $LogPath -Append
    }
    
    # Add a paragraph after the table
    "Adding paragraph after table" | Out-File -FilePath $LogPath -Append
    $Selection.TypeParagraph()
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
    
    # Skip table creation if there are no headers or rows
    if ($columnCount -eq 0 -or $rowCount -eq 0) {
        return
    }
    
    # Create a table with headers + data rows
    $table = $Selection.Tables.Add($Selection.Range, $rowCount + 1, $columnCount)
    
    # Set table properties for better appearance
    $table.Borders.Enable = $true
    
    # Apply a clean, professional table style without colored background
    $table.Style = "Table Grid"
    
    # Ensure no shading/background color
    for ($row = 1; $row -le $rowCount + 1; $row++) {
        for ($col = 1; $col -le $columnCount; $col++) {
            $table.Cell($row, $col).Shading.BackgroundPatternColor = -16777216 # wdColorAutomatic
        }
    }
    
    # Add headers
    for ($col = 0; $col -lt $columnCount; $col++) {
        if ($col -lt $Headers.Count) {
            $table.Cell(1, $col + 1).Range.Text = $Headers[$col]
        }
    }
    
    # Make the header row bold
    $table.Rows.Item(1).Range.Bold = $true
    
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
    
    # Auto-fit contents
    $table.AutoFitBehavior(1) # wdAutoFitContent
    
    # Add a paragraph after the table
    $Selection.TypeParagraph()
}

# Execute the script
Process-Document -InputFile $InputFile -OutputFile $OutputFile -ShowWord:$ShowWord 
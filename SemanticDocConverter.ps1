param (
    [Parameter(Mandatory = $true)]
    [string]$InputFile,
    
    [Parameter(Mandatory = $true)]
    [string]$OutputFile,
    
    [Parameter(Mandatory = $false)]
    [switch]$ShowWord,
    
    [Parameter(Mandatory = $false)]
    [switch]$Debug
)

function Write-DebugLog {
    param ([string]$Message)
    
    if ($Debug) {
        Write-Host "DEBUG: $Message" -ForegroundColor Cyan
    }
}

function Analyze-DocumentStructure {
    param ([string[]]$Lines)
    
    Write-DebugLog "Analyzing document structure..."
    
    $structure = @()
    $i = 0
    
    while ($i -lt $Lines.Count) {
        $line = $Lines[$i]
        $element = @{
            Type = "Unknown"
            StartLine = $i
            EndLine = $i
            Content = @($line)
            HeadingLevel = 0
            TableColumns = @()
        }
        
        # Detect heading
        if ($line -match '^(#+)\s+(.+)$') {
            $headingLevel = $matches[1].Length
            $headingText = $matches[2]
            
            $element.Type = "Heading"
            $element.HeadingLevel = $headingLevel
            $element.Content = @($headingText)
            
            Write-DebugLog "Found heading level $headingLevel at line $i"
            $structure += $element
            $i++
            continue
        }
        
        # Detect real table (must have multiple pipe characters and data rows)
        if ($line -match '\|.*\|.*\|') {
            Write-DebugLog "Potential table starting at line $i"
            
            # Collect all lines that could be part of the table
            $tableLines = @($line)
            $lineIndex = $i + 1
            
            while ($lineIndex -lt $Lines.Count -and 
                   $Lines[$lineIndex] -match '\|' -and 
                   -not [string]::IsNullOrWhiteSpace($Lines[$lineIndex]) -and
                   -not ($Lines[$lineIndex] -match '^#+\s+')) {
                
                $tableLines += $Lines[$lineIndex]
                $lineIndex++
            }
            
            # Analyze if this is a real table (needs at least header + separator + one data row)
            if ($tableLines.Count -ge 3 -and 
                $tableLines[1] -match '^[\|\s\-:]+$') {
                
                # This is a genuine table
                $columnHeaders = Extract-TableHeaders $tableLines[0]
                
                $element.Type = "Table"
                $element.Content = $tableLines
                $element.EndLine = $i + $tableLines.Count - 1
                $element.TableColumns = $columnHeaders
                
                Write-DebugLog "Confirmed genuine table ($($columnHeaders.Count) columns) from line $i to $($element.EndLine)"
                $structure += $element
                
                $i = $lineIndex
                continue
            }
        }
        
        # Detect paragraph (consecutive non-empty lines)
        if (-not [string]::IsNullOrWhiteSpace($line)) {
            $paragraphLines = @($line)
            $lineIndex = $i + 1
            
            while ($lineIndex -lt $Lines.Count -and 
                   -not [string]::IsNullOrWhiteSpace($Lines[$lineIndex]) -and
                   -not ($Lines[$lineIndex] -match '^#+\s+') -and
                   -not ($Lines[$lineIndex] -match '\|.*\|.*\|')) {
                
                $paragraphLines += $Lines[$lineIndex]
                $lineIndex++
            }
            
            $element.Type = "Paragraph"
            $element.Content = $paragraphLines
            $element.EndLine = $i + $paragraphLines.Count - 1
            
            Write-DebugLog "Found paragraph from line $i to $($element.EndLine)"
            $structure += $element
            
            $i = $lineIndex
            continue
        }
        
        # Empty line
        if ([string]::IsNullOrWhiteSpace($line)) {
            $element.Type = "EmptyLine"
            
            Write-DebugLog "Found empty line at line $i"
            $structure += $element
            $i++
            continue
        }
        
        # If we reach here, just treat as plain text
        $element.Type = "Text"
        
        Write-DebugLog "Found plain text at line $i"
        $structure += $element
        $i++
    }
    
    return $structure
}

function Extract-TableHeaders {
    param ([string]$HeaderLine)
    
    $headerLine = $HeaderLine.Trim()
    if ($headerLine.StartsWith('|')) {
        $headerLine = $headerLine.Substring(1)
    }
    if ($headerLine.EndsWith('|')) {
        $headerLine = $headerLine.Substring(0, $headerLine.Length - 1)
    }
    
    $headers = $headerLine.Split('|') | ForEach-Object { $_.Trim() }
    return $headers
}

function Format-Document {
    param (
        $Word,
        $Selection,
        $Structure
    )
    
    $elementCount = $Structure.Count
    Write-DebugLog "Formatting document with $elementCount elements..."
    
    foreach ($element in $Structure) {
        switch ($element.Type) {
            "Heading" {
                $headingLevel = $element.HeadingLevel
                $headingText = $element.Content[0]
                
                Write-DebugLog "Processing heading level $headingLevel - $headingText"
                
                $Selection.Style = "Heading $headingLevel"
                $Selection.TypeText($headingText)
                $Selection.TypeParagraph()
            }
            
            "Table" {
                Write-DebugLog "Processing table with $($element.TableColumns.Count) columns"
                Process-Table $Word $Selection $element
            }
            
            "Paragraph" {
                Write-DebugLog "Processing paragraph with $($element.Content.Count) lines"
                
                foreach ($line in $element.Content) {
                    $Selection.TypeText($line)
                    $Selection.TypeParagraph()
                }
            }
            
            "EmptyLine" {
                Write-DebugLog "Processing empty line"
                $Selection.TypeParagraph()
            }
            
            "Text" {
                Write-DebugLog "Processing text line: $($element.Content[0])"
                $Selection.TypeText($element.Content[0])
                $Selection.TypeParagraph()
            }
        }
    }
}

function Process-Table {
    param (
        $Word,
        $Selection,
        $TableElement
    )
    
    $tableLines = $TableElement.Content
    
    # Skip if not enough lines for a valid table
    if ($tableLines.Count -lt 3) {
        Write-DebugLog "Skipping table - not enough lines"
        foreach ($line in $tableLines) {
            $Selection.TypeText($line)
            $Selection.TypeParagraph()
        }
        return
    }
    
    # Extract header row
    $headerRow = $tableLines[0]
    $headers = Extract-TableHeaders $headerRow
    
    # Skip separator row
    # Process data rows
    $dataRows = @()
    for ($i = 2; $i -lt $tableLines.Count; $i++) {
        $row = $tableLines[$i]
        
        # Skip separator rows
        if ($row -match '^\s*[\|\-\+:]+\s*$') {
            continue
        }
        
        # Process data row
        $row = $row.Trim()
        if ($row.StartsWith('|')) {
            $row = $row.Substring(1)
        }
        if ($row.EndsWith('|')) {
            $row = $row.Substring(0, $row.Length - 1)
        }
        
        $cells = $row.Split('|') | ForEach-Object { $_.Trim() }
        $dataRows += , $cells
    }
    
    # Determine the real number of rows and columns
    $rowCount = 1 + $dataRows.Count # Header + data rows
    $colCount = $headers.Count
    
    # Only create a table if we have headers and data
    if ($rowCount -lt 2 -or $colCount -lt 1) {
        Write-DebugLog "Skipping table - not enough data"
        foreach ($line in $tableLines) {
            $Selection.TypeText($line)
            $Selection.TypeParagraph()
        }
        return
    }
    
    Write-DebugLog "Creating table with $rowCount rows and $colCount columns"
    
    # Create the table
    $table = $Selection.Tables.Add($Selection.Range, $rowCount, $colCount)
    $table.Borders.Enable = $true
    $table.Style = "Table Grid"
    
    # Add headers to first row
    for ($col = 0; $col -lt $headers.Count; $col++) {
        $table.Cell(1, $col + 1).Range.Text = $headers[$col]
    }
    
    # Bold the header row
    $table.Rows.Item(1).Range.Bold = $true
    
    # Add data rows
    for ($row = 0; $row -lt $dataRows.Count; $row++) {
        $cells = $dataRows[$row]
        
        for ($col = 0; $col -lt [Math]::Min($cells.Count, $colCount); $col++) {
            if ($col -lt $cells.Count) {
                $table.Cell($row + 2, $col + 1).Range.Text = $cells[$col]
            }
        }
    }
    
    # Auto-fit the table
    $table.AutoFitBehavior(1) # wdAutoFitContent
    
    # Add space after table
    $Selection.TypeParagraph()
}

# Main script execution
try {
    Write-DebugLog "Starting semantic document conversion..."
    
    # Check if input file exists
    if (-not (Test-Path $InputFile)) {
        throw "Input file not found: $InputFile"
    }
    
    # Read the file content
    $content = Get-Content -Path $InputFile -Raw
    $lines = $content -split "`r`n|\r|\n"
    
    Write-DebugLog "Read $($lines.Count) lines from input file"
    
    # First pass: Analyze document structure
    $documentStructure = Analyze-DocumentStructure $lines
    
    Write-DebugLog "Document analysis found $($documentStructure.Count) elements"
    
    # Create Word application
    $word = New-Object -ComObject Word.Application
    $word.Visible = [bool]$ShowWord
    
    # Create a new document
    $doc = $word.Documents.Add()
    $selection = $word.Selection
    
    # Second pass: Format document in Word
    Format-Document $word $selection $documentStructure
    
    # Save and close
    $absolutePath = [System.IO.Path]::GetFullPath($OutputFile)
    Write-DebugLog "Saving document to: $absolutePath"
    
    $doc.SaveAs([ref]$absolutePath, [ref]16) # 16 = wdFormatDocumentDefault
    $doc.Close()
    $word.Quit()
    
    Write-Host "Document converted successfully: $OutputFile" -ForegroundColor Green
}
catch {
    Write-Host "ERROR: $_" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    
    # Clean up
    if ($doc -ne $null) { try { $doc.Close($false) } catch { } }
    if ($word -ne $null) { try { $word.Quit() } catch { } }
}
finally {
    # Release COM objects
    if ($selection -ne $null) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($selection) | Out-Null }
    if ($doc -ne $null) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null }
    if ($word -ne $null) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
} 
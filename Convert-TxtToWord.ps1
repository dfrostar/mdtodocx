[CmdletBinding(SupportsShouldProcess=$true)]
param (
    [Parameter(Mandatory = $true)]
    [string]$InputFile,
    
    [Parameter(Mandatory = $true)]
    [string]$OutputFile,
    
    [Parameter(Mandatory = $false)]
    [switch]$ShowWord
)

# Create a log file
$logFile = "convert_log.txt"
"Starting conversion at $(Get-Date)" | Out-File -FilePath $logFile

function Log-Message {
    param([string]$Message)
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Out-File -FilePath $logFile -Append
    
    if ($VerbosePreference -eq 'Continue') {
        Write-Verbose $Message
    }
}

function Process-TableContent {
    param (
        $Word,
        $Selection,
        [array]$TableLines
    )
    
    # Need at least one row
    if ($TableLines.Count -lt 1) {
        Log-Message "No table lines found"
        return
    }
    
    Log-Message "Processing table with $($TableLines.Count) lines"
    
    # First, identify headers by parsing the first line with pipes
    $headerLine = $TableLines[0]
    $headers = @()
    
    # Extract header texts between pipes
    if ($headerLine -match '\|') {
        Log-Message "Header line has pipes: $headerLine"
        $parts = $headerLine.Split('|', [System.StringSplitOptions]::RemoveEmptyEntries)
        foreach ($part in $parts) {
            $headers += $part.Trim()
            Log-Message "  Added header: '$($part.Trim())'"
        }
    }
    else {
        # If no pipes in first line, use the first line as a single header
        $headers += $headerLine.Trim()
        Log-Message "Using single header: $($headerLine.Trim())"
    }
    
    # Remove any empty headers and ensure we have at least one
    $headers = $headers | Where-Object { $_ -ne "" }
    if ($headers.Count -eq 0) {
        $headers += "Column1"
        Log-Message "No valid headers found, using default 'Column1'"
    }
    
    Log-Message "Final headers count: $($headers.Count)"
    
    # Collect data rows
    $dataRows = @()
    for ($i = 1; $i -lt $TableLines.Count; $i++) {
        $line = $TableLines[$i]
        Log-Message "Processing potential data row: $line"
        
        # Skip separator rows (lines with only dashes, pipes, and plus signs)
        if ($line -match '^\s*[\|\+\-\s]+\s*$') {
            Log-Message "  Skipping separator row"
            continue
        }
        
        $rowData = @()
        if ($line -match '\|') {
            # Split by pipe and process each cell
            $parts = $line.Split('|', [System.StringSplitOptions]::RemoveEmptyEntries)
            Log-Message "  Line has pipes, found $($parts.Count) parts"
            foreach ($part in $parts) {
                $rowData += $part.Trim()
            }
        }
        else {
            # If no pipes, treat as single column
            $rowData += $line.Trim()
            Log-Message "  Line has no pipes, using as single cell"
        }
        
        # Ensure row has values
        if ($rowData.Count -gt 0) {
            $dataRows += ,$rowData
            Log-Message "  Added row with $($rowData.Count) cells"
        }
    }
    
    Log-Message "Total data rows collected: $($dataRows.Count)"
    
    # Determine the number of columns (use maximum of headers or any data row)
    $columnCount = $headers.Count
    foreach ($row in $dataRows) {
        if ($row.Count -gt $columnCount) {
            $columnCount = $row.Count
        }
    }
    
    Log-Message "Final column count for table: $columnCount"
    
    try {
        # Create the table (header row + data rows)
        $rowCount = 1 + $dataRows.Count
        Log-Message "Creating table with $rowCount rows and $columnCount columns"
        
        $table = $Selection.Tables.Add($Selection.Range, $rowCount, $columnCount)
        $table.Borders.Enable = $true
        $table.Style = "Table Grid"
        
        # Add headers to first row
        for ($col = 0; $col -lt $headers.Count -and $col -lt $columnCount; $col++) {
            $table.Cell(1, $col + 1).Range.Text = $headers[$col]
            Log-Message "  Added header to column $($col+1): '$($headers[$col])'"
        }
        
        # Bold the header row
        $table.Rows.Item(1).Range.Bold = $true
        
        # Add data rows
        for ($row = 0; $row -lt $dataRows.Count; $row++) {
            $rowData = $dataRows[$row]
            for ($col = 0; $col -lt $rowData.Count -and $col -lt $columnCount; $col++) {
                $table.Cell($row + 2, $col + 1).Range.Text = $rowData[$col]
            }
            Log-Message "  Added data row $($row+1) with $($rowData.Count) cells"
        }
        
        # AutoFit the table to content
        $table.AutoFitBehavior(1) # wdAutoFitContent = 1
        
        # Add space after table
        $Selection.TypeParagraph()
        Log-Message "Table creation complete"
    }
    catch {
        Log-Message "ERROR creating table: $_"
        Log-Message $_.ScriptStackTrace
    }
}

try {
    # Check if file exists
    if (-not (Test-Path $InputFile)) {
        $errorMsg = "Input file not found: $InputFile"
        Log-Message $errorMsg
        Write-Error $errorMsg
        exit 1
    }
    
    Log-Message "Reading input file: $InputFile"
    # Read file content
    $content = Get-Content -Path $InputFile -Raw
    $lines = $content -split "`r`n|\r|\n"
    Log-Message "Read $($lines.Count) lines from input file"
    
    # Create Word application
    Log-Message "Creating Word application"
    $word = New-Object -ComObject Word.Application
    $word.Visible = [bool]$ShowWord
    
    # Create new document
    $doc = $word.Documents.Add()
    $selection = $word.Selection
    Log-Message "Created new Word document"
    
    # Process each line
    $i = 0
    while ($i -lt $lines.Count) {
        $line = $lines[$i]
        Log-Message "Processing line $($i+1): $line"
        
        # Process headings (# or ## style)
        if ($line -match "^#+\s+(.+)$") {
            $headingLevel = ($line -split ' ')[0].Length
            $headingText = $matches[1]
            
            Log-Message "Found heading level $headingLevel`: $headingText"
            
            # Apply heading style
            $selection.Style = $doc.Styles.Item("Heading $headingLevel")
            $selection.TypeText($headingText)
            $selection.TypeParagraph()
            $i++
            continue
        }
        
        # Identify potential table sections
        # A table section starts with a line containing a pipe, or a line that looks like a header
        # followed by pipes or separator characters on the next line
        elseif ($line -match "\|" -or 
               ($i + 1 -lt $lines.Count -and $lines[$i+1] -match "[\|\-\+]")) {
            
            Log-Message "Potential table starting at line $($i+1)"
            
            # Start collecting table lines
            $tableLines = New-Object System.Collections.ArrayList
            $tableLines.Add($line) | Out-Null
            $i++
            
            # Continue collecting lines until we hit an empty line or a heading
            while ($i -lt $lines.Count -and 
                  -not [string]::IsNullOrWhiteSpace($lines[$i]) -and
                  -not ($lines[$i] -match "^#+\s")) {
                $tableLines.Add($lines[$i]) | Out-Null
                Log-Message "  Added table line: $($lines[$i])"
                $i++
            }
            
            # Process the collected table
            if ($tableLines.Count -gt 0) {
                Log-Message "Processing collected table with $($tableLines.Count) lines"
                Process-TableContent -Word $word -Selection $selection -TableLines $tableLines
            }
            
            continue
        }
        
        # Process regular text
        elseif (-not [string]::IsNullOrWhiteSpace($line)) {
            Log-Message "Regular text line: $line"
            $selection.TypeText($line)
            $selection.TypeParagraph()
            $i++
            continue
        }
        # Skip empty lines with paragraph
        else {
            Log-Message "Empty line, adding paragraph"
            $selection.TypeParagraph()
            $i++
            continue
        }
    }
    
    # Save and close
    $absolutePath = [System.IO.Path]::GetFullPath($OutputFile)
    Log-Message "Saving document to: $absolutePath"
    $doc.SaveAs([ref]$absolutePath, [ref]16) # 16 = wdFormatDocumentDefault
    $doc.Close()
    $word.Quit()
    
    Log-Message "Document converted successfully: $OutputFile"
}
catch {
    $errorMsg = "ERROR converting document: $_"
    Log-Message $errorMsg
    Log-Message $_.ScriptStackTrace
    Write-Error $errorMsg
    
    # Cleanup
    if ($doc -ne $null) { try { $doc.Close($false) } catch { Log-Message "Error closing document: $_" } }
    if ($word -ne $null) { try { $word.Quit() } catch { Log-Message "Error quitting Word: $_" } }
}
finally {
    # Release COM objects
    Log-Message "Releasing COM objects"
    if ($selection -ne $null) { 
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($selection) | Out-Null }
        catch { Log-Message "Error releasing selection object: $_" }
    }
    if ($doc -ne $null) { 
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null }
        catch { Log-Message "Error releasing document object: $_" }
    }
    if ($word -ne $null) { 
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null }
        catch { Log-Message "Error releasing Word application object: $_" }
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    Log-Message "Conversion process complete"
}

# Run the conversion
# Convert-TxtToWord.ps1 -InputFile $InputFile -OutputFile $OutputFile -ShowWord:$ShowWord 
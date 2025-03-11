param (
    [Parameter(Mandatory = $true)]
    [string]$InputFile,
    
    [Parameter(Mandatory = $true)]
    [string]$OutputFile,
    
    [Parameter(Mandatory = $false)]
    [switch]$ShowWord
)

try {
    # Check if file exists
    if (-not (Test-Path $InputFile)) {
        Write-Error "Input file not found: $InputFile"
        exit 1
    }
    
    # Read the file content
    $content = Get-Content -Path $InputFile -Raw
    $lines = $content -split "`r`n|\r|\n"
    
    # Create Word application
    $word = New-Object -ComObject Word.Application
    $word.Visible = [bool]$ShowWord
    
    # Create a new document
    $doc = $word.Documents.Add()
    $selection = $word.Selection
    
    # Process each line
    $i = 0
    while ($i -lt $lines.Count) {
        $line = $lines[$i]
        
        # Process headings (# or ## style)
        if ($line -match "^#+\s+(.+)$") {
            $headingLevel = ($line -split ' ')[0].Length
            $headingText = $matches[1]
            
            # Apply heading style
            $selection.Style = "Heading $headingLevel"
            $selection.TypeText($headingText)
            $selection.TypeParagraph()
            $i++
            continue
        }
        
        # Process tables (lines with pipe characters)
        elseif ($line -match "\|") {
            # Collect table lines
            $tableLines = New-Object System.Collections.ArrayList
            
            # Continue collecting lines until we find a line without a pipe or reach the end
            while ($i -lt $lines.Count -and $lines[$i] -match "\|") {
                $currentLine = $lines[$i]
                
                # Skip separator lines (e.g., |-----|-----|)
                if (-not ($currentLine -match "^[\|\s\-\+]+$")) {
                    $tableLines.Add($currentLine) | Out-Null
                }
                
                $i++
            }
            
            # Process table
            if ($tableLines.Count -gt 0) {
                # Determine the number of columns
                $maxColumns = 0
                foreach ($tableLine in $tableLines) {
                    $pipeCount = ($tableLine.ToCharArray() | Where-Object { $_ -eq '|' } | Measure-Object).Count
                    $colCount = $pipeCount - 1
                    if ($colCount -gt $maxColumns) {
                        $maxColumns = $colCount
                    }
                }
                
                # Create table
                $table = $selection.Tables.Add($selection.Range, $tableLines.Count, $maxColumns)
                $table.Borders.Enable = $true
                $table.Style = "Table Grid"
                
                # Fill in table data
                for ($rowIndex = 0; $rowIndex -lt $tableLines.Count; $rowIndex++) {
                    $rowLine = $tableLines[$rowIndex]
                    
                    # Split the line by pipes
                    $columnTexts = $rowLine.Split('|', [System.StringSplitOptions]::RemoveEmptyEntries)
                    
                    # Add each cell
                    for ($colIndex = 0; $colIndex -lt $columnTexts.Count -and $colIndex -lt $maxColumns; $colIndex++) {
                        $cellText = $columnTexts[$colIndex].Trim()
                        $table.Cell($rowIndex + 1, $colIndex + 1).Range.Text = $cellText
                    }
                }
                
                # Bold the header row
                if ($tableLines.Count -gt 0) {
                    $table.Rows.Item(1).Range.Bold = $true
                }
                
                # Auto-fit the table
                $table.AutoFitBehavior(1) # wdAutoFitContent
                
                # Add a paragraph after the table
                $selection.TypeParagraph()
            }
            
            continue
        }
        
        # Process regular text
        else {
            if (-not [string]::IsNullOrWhiteSpace($line)) {
                $selection.TypeText($line)
            }
            
            $selection.TypeParagraph()
            $i++
        }
    }
    
    # Save and close
    $doc.SaveAs([ref]([System.IO.Path]::GetFullPath($OutputFile)), [ref]16) # wdFormatDocumentDefault
    $doc.Close()
    $word.Quit()
    
    Write-Host "Document converted successfully: $OutputFile"
}
catch {
    Write-Error "Error: $_"
    Write-Host $_.ScriptStackTrace
    
    # Clean up
    if ($doc -ne $null) { 
        try { $doc.Close($false) } catch { }
    }
    if ($word -ne $null) { 
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
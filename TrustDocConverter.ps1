[CmdletBinding(SupportsShouldProcess=$true)]
param (
    [Parameter(Mandatory = $true)]
    [string]$InputFile,
    
    [Parameter(Mandatory = $true)]
    [string]$OutputFile,
    
    [Parameter(Mandatory = $false)]
    [switch]$ShowWord
)

function Write-DebugLog {
    param ([string]$Message)
    
    if ($VerbosePreference -eq 'Continue') {
        Write-Verbose $Message
    }
}

function Is-SectionHeading {
    param ([string]$Line)
    
    # Check for common section headers in trust documents
    $sectionPatterns = @(
        "Real Estate Properties",
        "Financial Accounts",
        "Retirement Accounts",
        "Vehicles",
        "Life Insurance Policies",
        "Business Interests",
        "Valuable Personal Property",
        "Intellectual Property",
        "Transfer Status Tracking",
        "Professional Contact",
        "Important Dates"
    )
    
    foreach ($pattern in $sectionPatterns) {
        if ($Line -match $pattern) {
            return $true
        }
    }
    
    return $false
}

function Is-TableData {
    param ([string]$Line)
    
    # Real tables in trust documents typically have:
    # 1. Multiple pipe characters (at least 3)
    # 2. Consistent structure
    
    if ($Line -match '\|.*\|.*\|') {
        # Count pipe characters (should be at least 3 for a real table)
        $pipeCount = ($Line.ToCharArray() | Where-Object { $_ -eq '|' } | Measure-Object).Count
        
        if ($pipeCount -ge 3) {
            return $true
        }
    }
    
    return $false
}

function Is-Separator {
    param ([string]$Line)
    
    # Separators are lines with only dashes, pipes, and whitespace
    return $Line -match '^\s*[\|\-\+:]+\s*$'
}

function Analyze-TrustDocument {
    param ([string[]]$Lines)
    
    Write-DebugLog "Analyzing trust document structure..."
    
    $structure = @()
    $i = 0
    $currentSection = "General"
    
    while ($i -lt $Lines.Count) {
        $line = $Lines[$i]
        $element = @{
            Type = "Unknown"
            StartLine = $i
            EndLine = $i
            Content = @($line)
            HeadingLevel = 0
            Section = $currentSection
            TableColumns = @()
        }
        
        # Detect main heading (# or ## style)
        if ($line -match '^(#+)\s+(.+)$') {
            $headingLevel = $matches[1].Length
            $headingText = $matches[2]
            
            $element.Type = "Heading"
            $element.HeadingLevel = $headingLevel
            $element.Content = @($headingText)
            
            # Update current section for level 2 or 3 headings
            if ($headingLevel -in @(2, 3)) {
                $currentSection = $headingText
            }
            
            Write-DebugLog "Found heading level $headingLevel at line $i"
            $structure += $element
            $i++
            continue
        }
        
        # Detect section heading (not with # but known section name)
        elseif (Is-SectionHeading $line) {
            $element.Type = "SectionHeading"
            $element.Content = @($line)
            $currentSection = $line
            
            Write-DebugLog "Found section heading at line $i - $line"
            $structure += $element
            $i++
            continue
        }
        
        # Detect table structure (realistic tables, not just any line with a pipe)
        elseif (Is-TableData $line) {
            Write-DebugLog "Potential table starting at line $i"
            
            # Collect all lines that could be part of this table
            $tableLines = @($line)
            $lineIndex = $i + 1
            $foundSeparator = $false
            
            # Look ahead for more table lines
            while ($lineIndex -lt $Lines.Count) {
                $nextLine = $Lines[$lineIndex]
                
                # Stop if we hit a heading
                if ($nextLine -match '^#+\s+' -or (Is-SectionHeading $nextLine)) {
                    break
                }
                
                # Check for separator line after header
                if ($lineIndex -eq $i + 1 -and (Is-Separator $nextLine)) {
                    $foundSeparator = $true
                }
                
                # Add if it looks like table data
                if ((Is-TableData $nextLine) -or (Is-Separator $nextLine)) {
                    $tableLines += $nextLine
                    $lineIndex++
                    continue
                }
                
                # Empty line ends table only if we're already past the first few rows
                if ([string]::IsNullOrWhiteSpace($nextLine) -and $tableLines.Count -gt 2) {
                    break
                }
                
                # Non-table line ends collection
                if (-not [string]::IsNullOrWhiteSpace($nextLine)) {
                    break
                }
                
                # Empty line (but continue collecting)
                $lineIndex++
            }
            
            # Only process as table if we have a proper structure
            if ($tableLines.Count -ge 2 -and ($foundSeparator -or $tableLines.Count -ge 3)) {
                $element.Type = "Table"
                $element.Content = $tableLines
                $element.EndLine = $i + $tableLines.Count - 1
                $element.Section = $currentSection
                
                Write-DebugLog "Confirmed table from line $i to $($element.EndLine) in section $currentSection"
                $structure += $element
                
                $i = $lineIndex
                continue
            }
            else {
                # Not a real table, just process as text
                $element.Type = "Text"
                Write-DebugLog "Line with pipes treated as text: $line"
                $structure += $element
                $i++
                continue
            }
        }
        
        # Detect paragraph (consecutive non-empty lines)
        elseif (-not [string]::IsNullOrWhiteSpace($line)) {
            $paragraphLines = @($line)
            $lineIndex = $i + 1
            
            while ($lineIndex -lt $Lines.Count -and 
                   -not [string]::IsNullOrWhiteSpace($Lines[$lineIndex]) -and
                   -not ($Lines[$lineIndex] -match '^#+\s+') -and
                   -not (Is-SectionHeading $Lines[$lineIndex]) -and
                   -not (Is-TableData $Lines[$lineIndex])) {
                
                $paragraphLines += $Lines[$lineIndex]
                $lineIndex++
            }
            
            $element.Type = "Paragraph"
            $element.Content = $paragraphLines
            $element.EndLine = $i + $paragraphLines.Count - 1
            
            Write-DebugLog "Found paragraph from line $i to $($element.EndLine) in section $currentSection"
            $structure += $element
            
            $i = $lineIndex
            continue
        }
        
        # Empty line
        elseif ([string]::IsNullOrWhiteSpace($line)) {
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

function Format-TrustDocument {
    param (
        $Word,
        $Selection,
        $Structure
    )
    
    $elementCount = $Structure.Count
    Write-DebugLog "Formatting trust document with $elementCount elements..."
    
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
            
            "SectionHeading" {
                $sectionText = $element.Content[0]
                
                Write-DebugLog "Processing section heading: $sectionText"
                
                $Selection.Style = "Heading 3"
                $Selection.TypeText($sectionText)
                $Selection.TypeParagraph()
            }
            
            "Table" {
                Write-DebugLog "Processing table in section $($element.Section)"
                Process-TrustTable $Word $Selection $element
            }
            
            "Paragraph" {
                Write-DebugLog "Processing paragraph with $($element.Content.Count) lines"
                
                foreach ($line in $element.Content) {
                    $Selection.TypeText($line)
                    $Selection.TypeParagraph()
                }
            }
            
            "EmptyLine" {
                $Selection.TypeParagraph()
            }
            
            "Text" {
                $Selection.TypeText($element.Content[0])
                $Selection.TypeParagraph()
            }
        }
    }
}

function Process-TrustTable {
    param (
        $Word,
        $Selection,
        $TableElement
    )
    
    $tableLines = $TableElement.Content
    
    # Skip if not enough lines for a valid table
    if ($tableLines.Count -lt 2) {
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
    
    if ($headers.Count -eq 0) {
        Write-DebugLog "No valid headers found, skipping table"
        foreach ($line in $tableLines) {
            $Selection.TypeText($line)
            $Selection.TypeParagraph()
        }
        return
    }
    
    # Process data rows (skip separator rows)
    $dataRows = @()
    for ($i = 1; $i -lt $tableLines.Count; $i++) {
        $row = $tableLines[$i]
        
        # Skip separator rows
        if (Is-Separator $row) {
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
        
        # Skip empty rows
        if ([string]::IsNullOrWhiteSpace($row)) {
            continue
        }
        
        $cells = $row.Split('|') | ForEach-Object { $_.Trim() }
        
        # Skip rows with empty cells
        if ($cells.Count -gt 0) {
            $allEmpty = $true
            foreach ($cell in $cells) {
                if (-not [string]::IsNullOrWhiteSpace($cell)) {
                    $allEmpty = $false
                    break
                }
            }
            
            if (-not $allEmpty) {
                $dataRows += , $cells
            }
        }
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
        $headerText = $headers[$col].Trim()
        if (-not [string]::IsNullOrWhiteSpace($headerText)) {
            $table.Cell(1, $col + 1).Range.Text = $headerText
        }
    }
    
    # Bold the header row
    $table.Rows.Item(1).Range.Bold = $true
    
    # Add data rows
    for ($row = 0; $row -lt $dataRows.Count; $row++) {
        $cells = $dataRows[$row]
        
        for ($col = 0; $col -lt [Math]::Min($cells.Count, $colCount); $col++) {
            if ($col -lt $cells.Count) {
                $cellText = $cells[$col].Trim()
                if (-not [string]::IsNullOrWhiteSpace($cellText)) {
                    $table.Cell($row + 2, $col + 1).Range.Text = $cellText
                }
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
    Write-DebugLog "Starting trust document conversion..."
    
    # Check if input file exists
    if (-not (Test-Path $InputFile)) {
        throw "Input file not found: $InputFile"
    }
    
    # Read the file content
    $content = Get-Content -Path $InputFile -Raw
    $lines = $content -split "`r`n|\r|\n"
    
    Write-DebugLog "Read $($lines.Count) lines from input file"
    
    # First pass: Analyze document structure
    $documentStructure = Analyze-TrustDocument $lines
    
    Write-DebugLog "Document analysis found $($documentStructure.Count) elements"
    
    # Create Word application
    $word = New-Object -ComObject Word.Application
    $word.Visible = [bool]$ShowWord
    
    # Create a new document
    $doc = $word.Documents.Add()
    $selection = $word.Selection
    
    # Second pass: Format document in Word
    Format-TrustDocument $word $selection $documentStructure
    
    # Save and close
    $absolutePath = [System.IO.Path]::GetFullPath($OutputFile)
    Write-DebugLog "Saving document to: $absolutePath"
    
    $doc.SaveAs([ref]$absolutePath, [ref]16) # 16 = wdFormatDocumentDefault
    $doc.Close()
    $word.Quit()
    
    Write-Host "Trust document converted successfully: $OutputFile" -ForegroundColor Green
    Write-Host "This specialized converter provides better understanding of trust inventory structure." -ForegroundColor Green
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
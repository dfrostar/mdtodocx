# Test script for TrustInventory document conversion
param (
    [Parameter(Mandatory = $false)]
    [string]$InputFile = "trust-asset-inventory.txt",
    
    [Parameter(Mandatory = $false)]
    [string]$OutputFile = "trust-asset-inventory-test.docx",
    
    [Parameter(Mandatory = $false)]
    [switch]$ShowWord
)

# Load utility functions
. ".\DocConverter\Private\Common-Utils.ps1"

# Helper function to check if a line is a section heading
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

# Helper function to check if a line is table data
function Is-TableData {
    param ([string]$Line)
    
    # Real tables in trust documents typically have multiple pipe characters
    if ($Line -match '\|.*\|.*\|') {
        $pipeCount = ($Line.ToCharArray() | Where-Object { $_ -eq '|' } | Measure-Object).Count
        if ($pipeCount -ge 3) {
            return $true
        }
    }
    
    return $false
}

# Helper function to check if a line is a separator
function Is-Separator {
    param ([string]$Line)
    
    # Separators are lines with only dashes, pipes, and whitespace
    return $Line -match '^\s*[\|\-\+:]+\s*$'
}

# Initialize Word application function
function Initialize-WordApplication {
    param (
        [Parameter(Mandatory = $false)]
        [switch]$ShowWord
    )
    
    Write-Host "Initializing Word application"
    $word = New-Object -ComObject Word.Application
    $word.Visible = [bool]$ShowWord
    
    return $word
}

# Close Word objects function
function Close-WordObjects {
    param (
        [Parameter(Mandatory = $false)]
        [System.Object]$Document,
        
        [Parameter(Mandatory = $false)]
        [System.Object]$WordApp,
        
        [Parameter(Mandatory = $false)]
        [System.Object]$Selection
    )
    
    Write-Host "Cleaning up Word objects"
    
    # Clean up
    if ($Document -ne $null) { 
        try { 
            $Document.Close($false)
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Document) | Out-Null 
        } 
        catch { } 
    }
    
    if ($Selection -ne $null) { 
        try { 
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Selection) | Out-Null 
        } 
        catch { } 
    }
    
    if ($WordApp -ne $null) { 
        try { 
            $WordApp.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($WordApp) | Out-Null 
        } 
        catch { } 
    }
    
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

# Extract table headers function
function Extract-TableHeaders {
    param ([string]$HeaderLine)
    
    Write-Host "Extracting table headers from: $HeaderLine"
    
    $headerLine = $HeaderLine.Trim()
    if ($headerLine.StartsWith('|')) {
        $headerLine = $headerLine.Substring(1)
    }
    if ($headerLine.EndsWith('|')) {
        $headerLine = $headerLine.Substring(0, $headerLine.Length - 1)
    }
    
    # Split by pipe character and trim each header
    $headers = $headerLine.Split('|') | ForEach-Object { $_.Trim() }
    
    # Filter out empty headers and ensure we have valid headers
    $headers = $headers | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    
    # If we have no valid headers, try to create placeholder headers
    if ($headers.Count -eq 0) {
        $pipeCount = ($HeaderLine.ToCharArray() | Where-Object { $_ -eq '|' } | Measure-Object).Count
        if ($pipeCount -gt 0) {
            $headers = @()
            for ($i = 1; $i -le ($pipeCount + 1); $i++) {
                $headers += "Column $i"
            }
        }
    }
    
    return $headers
}

# Function to analyze the document structure
function Analyze-TrustDocument {
    param ([string[]]$Lines)
    
    Write-Host "Analyzing trust document structure..."
    
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
            
            Write-Host "Found heading level $headingLevel at line $i"
            $structure += $element
            $i++
            continue
        }
        
        # Detect section heading (not with # but known section name)
        elseif (Is-SectionHeading $line) {
            $element.Type = "SectionHeading"
            $element.Content = @($line)
            $currentSection = $line
            
            Write-Host "Found section heading at line $i - $line"
            $structure += $element
            $i++
            continue
        }
        
        # Detect table structure (realistic tables, not just any line with a pipe)
        elseif (Is-TableData $line) {
            Write-Host "Potential table starting at line $i"
            
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
                
                Write-Host "Confirmed table from line $i to $($element.EndLine) in section $currentSection"
                $structure += $element
                
                $i = $lineIndex
                continue
            }
            else {
                # Not a real table, just process as text
                $element.Type = "Text"
                Write-Host "Line with pipes treated as text: $line"
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
            
            Write-Host "Found paragraph from line $i to $($element.EndLine) in section $currentSection"
            $structure += $element
            
            $i = $lineIndex
            continue
        }
        
        # Empty line
        elseif ([string]::IsNullOrWhiteSpace($line)) {
            $element.Type = "EmptyLine"
            
            Write-Host "Found empty line at line $i"
            $structure += $element
            $i++
            continue
        }
        
        # If we reach here, just treat as plain text
        $element.Type = "Text"
        
        Write-Host "Found plain text at line $i"
        $structure += $element
        $i++
    }
    
    return $structure
}

# Function to format the trust document in Word
function Format-TrustDocument {
    param (
        $Word,
        $Selection,
        $Structure
    )
    
    $elementCount = $Structure.Count
    Write-Host "Formatting trust document with $elementCount elements..."
    
    foreach ($element in $Structure) {
        switch ($element.Type) {
            "Heading" {
                $headingLevel = $element.HeadingLevel
                $headingText = $element.Content[0]
                
                Write-Host "Processing heading level $headingLevel - $headingText"
                
                $Selection.Style = "Heading $headingLevel"
                $Selection.TypeText($headingText)
                $Selection.TypeParagraph()
            }
            
            "SectionHeading" {
                $sectionText = $element.Content[0]
                
                Write-Host "Processing section heading: $sectionText"
                
                $Selection.Style = "Heading 3"
                $Selection.TypeText($sectionText)
                $Selection.TypeParagraph()
            }
            
            "Table" {
                Write-Host "Processing table in section $($element.Section)"
                
                $tableLines = $element.Content
                
                # Skip if not enough lines for a valid table
                if ($tableLines.Count -lt 2) {
                    Write-Host "Skipping table - not enough lines"
                    foreach ($line in $tableLines) {
                        $Selection.TypeText($line)
                        $Selection.TypeParagraph()
                    }
                    continue
                }
                
                # Extract header row
                $headerRow = $tableLines[0]
                $headers = Extract-TableHeaders $headerRow
                
                if ($headers.Count -eq 0) {
                    Write-Host "No valid headers found, skipping table"
                    foreach ($line in $tableLines) {
                        $Selection.TypeText($line)
                        $Selection.TypeParagraph()
                    }
                    continue
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
                    Write-Host "Skipping table - not enough data"
                    foreach ($line in $tableLines) {
                        $Selection.TypeText($line)
                        $Selection.TypeParagraph()
                    }
                    continue
                }
                
                Write-Host "Creating table with $rowCount rows and $colCount columns"
                
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
                
                # Set table properties for better appearance
                $table.AllowAutoFit = $true
                $table.AutoFitBehavior(1) # wdAutoFitContent = 1
                
                # Apply a cleaner table style if available
                try {
                    $table.Style = "Table Grid"
                    $table.ApplyStyleHeadingRows = $true
                    $table.ApplyStyleLastRow = $false
                    $table.ApplyStyleFirstColumn = $false
                    $table.ApplyStyleLastColumn = $false
                    $table.ApplyStyleRowBands = $true
                    $table.ApplyStyleColumnBands = $false
                } catch {
                    Write-Host "Could not apply all table styles"
                }
                
                # Add space after table
                $Selection.TypeParagraph()
            }
            
            "Paragraph" {
                Write-Host "Processing paragraph with $($element.Content.Count) lines"
                
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

# Main conversion process
try {
    Write-Host "Starting trust document conversion..."
    
    # Check if input file exists
    if (-not (Test-Path $InputFile)) {
        throw "Input file not found: $InputFile"
    }
    
    # Read the file content
    $content = Get-Content -Path $InputFile -Raw
    $lines = $content -split "`r`n|\r|\n"
    
    Write-Host "Read $($lines.Count) lines from input file"
    
    # First pass: Analyze document structure
    $documentStructure = Analyze-TrustDocument $lines
    
    Write-Host "Document analysis found $($documentStructure.Count) elements"
    
    # Initialize Word
    $word = Initialize-WordApplication -ShowWord:$ShowWord
    $doc = $word.Documents.Add()
    
    # Set document properties for better formatting
    $doc.PageSetup.LeftMargin = 72 # 1 inch in points
    $doc.PageSetup.RightMargin = 72
    $doc.PageSetup.TopMargin = 72
    $doc.PageSetup.BottomMargin = 72
    
    # Set default font
    $word.Selection.Font.Name = "Calibri"
    $word.Selection.Font.Size = 11
    
    $selection = $word.Selection
    
    # Second pass: Format document in Word
    Format-TrustDocument $word $selection $documentStructure
    
    # Save and close
    $absolutePath = [System.IO.Path]::GetFullPath($OutputFile)
    Write-Host "Saving document to: $absolutePath"
    
    $doc.SaveAs([ref]$absolutePath, [ref]16) # 16 = wdFormatDocumentDefault
    $doc.Close()
    $word.Quit()
    
    Write-Host "Trust document converted successfully: $OutputFile" -ForegroundColor Green
}
catch {
    Write-Host "ERROR: $_" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    
    # Clean up
    if ($doc -ne $null) { try { $doc.Close($false) } catch { } }
    if ($word -ne $null) { try { $word.Quit() } catch { } }
    
    throw
}
finally {
    # Release COM objects
    Close-WordObjects -Document $doc -WordApp $word -Selection $selection
} 
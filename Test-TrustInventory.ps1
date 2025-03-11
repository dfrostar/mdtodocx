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
        if ($pipeCount -ge 2) {
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

# Extract table headers function with improved parsing
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
    $validHeaders = @()
    foreach ($header in $headers) {
        if (-not [string]::IsNullOrWhiteSpace($header)) {
            # Clean up any markdown formatting symbols - fixed regex patterns
            $cleanHeader = $header.Replace("**", "").Replace("*", "").Replace("__", "").Replace("_", "")
            $validHeaders += $cleanHeader
        }
        else {
            # Use column placeholder for empty headers
            $validHeaders += "Column"
        }
    }
    
    return $validHeaders
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
                
                # Set heading styles with improved formatting
                $Selection.Style = "Heading $headingLevel"
                
                # Format headings based on level
                switch ($headingLevel) {
                    1 {
                        $Selection.Font.Size = 18
                        $Selection.Font.Bold = $true
                        $Selection.Font.Color = 5460735 # Dark blue
                    }
                    2 {
                        $Selection.Font.Size = 16
                        $Selection.Font.Bold = $true
                        $Selection.Font.Color = 6710932 # Dark teal
                    }
                    3 {
                        $Selection.Font.Size = 14
                        $Selection.Font.Bold = $true
                        $Selection.Font.Color = 10053222 # Dark gray
                    }
                    default {
                        $Selection.Font.Size = 12
                        $Selection.Font.Bold = $true
                    }
                }
                
                $Selection.TypeText($headingText)
                $Selection.TypeParagraph()
                
                # Add extra space after higher level headings
                if ($headingLevel -le 2) {
                    $Selection.TypeParagraph()
                }
            }
            
            "SectionHeading" {
                $sectionText = $element.Content[0]
                
                Write-Host "Processing section heading: $sectionText"
                
                $Selection.Style = "Heading 3"
                $Selection.Font.Size = 14
                $Selection.Font.Bold = $true
                $Selection.Font.Color = 10053222 # Dark gray
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
                
                # 1. Extract header row and count ACTUAL columns
                $headerRow = $null
                $dataRows = @()
                
                # Find the header row (first non-separator row)
                foreach ($line in $tableLines) {
                    if (-not (Is-Separator $line) -and (Is-TableData $line)) {
                        $headerRow = $line
                        break
                    }
                }
                
                if (-not $headerRow) {
                    Write-Host "No valid header row found, skipping table"
                    foreach ($line in $tableLines) {
                        $Selection.TypeText($line)
                        $Selection.TypeParagraph()
                    }
                    continue
                }
                
                # Extract and clean up headers
                $headers = Extract-TableHeaders $headerRow
                
                if ($headers.Count -eq 0) {
                    Write-Host "No valid headers found, skipping table"
                    foreach ($line in $tableLines) {
                        $Selection.TypeText($line)
                        $Selection.TypeParagraph()
                    }
                    continue
                }
                
                # 2. Process all data rows, ignoring separator rows
                foreach ($line in $tableLines) {
                    # Skip header row
                    if ($line -eq $headerRow) {
                        continue
                    }
                    
                    # Skip separator rows
                    if (Is-Separator $line) {
                        continue
                    }
                    
                    # Process data row if it contains pipes
                    if (Is-TableData $line) {
                        $row = $line.Trim()
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
                        
                        # Split the row into cells
                        $cells = $row.Split('|') | ForEach-Object { $_.Trim() }
                        
                        # Skip rows with all empty cells
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
                
                # 3. LIMIT max columns to 5 - this prevents the "value out of range" error
                $maxColumns = 5
                $originalColCount = $headers.Count
                $colCount = [Math]::Min($originalColCount, $maxColumns)
                
                # If we have more than maxColumns, we'll create multiple tables
                $tableCount = [Math]::Ceiling($originalColCount / $maxColumns)
                Write-Host "Table has $originalColCount columns, creating $tableCount table(s) with max $maxColumns columns each"
                
                for ($tableIndex = 0; $tableIndex -lt $tableCount; $tableIndex++) {
                    # Calculate the column range for this sub-table
                    $startCol = $tableIndex * $maxColumns
                    $endCol = [Math]::Min(($tableIndex + 1) * $maxColumns - 1, $originalColCount - 1)
                    $currentColCount = $endCol - $startCol + 1
                    
                    if ($tableIndex > 0) {
                        # Add separator between tables
                        $Selection.TypeParagraph()
                        
                        # Add a subtitle for continuation tables
                        $Selection.Font.Bold = $true
                        $Selection.Font.Size = 12
                        $Selection.TypeText("Table Continued (Columns $($startCol + 1) to $($endCol + 1))")
                        $Selection.Font.Bold = $false
                        $Selection.TypeParagraph()
                    }
                    
                    $rowCount = 1 + $dataRows.Count # Header + data rows
                    
                    Write-Host "Creating table $($tableIndex + 1) with $rowCount rows and $currentColCount columns"
                    
                    try {
                        # Create the table with exact dimensions
                        $table = $Selection.Tables.Add($Selection.Range, $rowCount, $currentColCount)
                        
                        # Basic table properties for cleaner appearance
                        $table.Borders.Enable = $true
                        $table.Borders.InsideLineStyle = 1 # wdLineStyleSingle  
                        $table.Borders.OutsideLineStyle = 1 # wdLineStyleSingle
                        
                        # Apply simple grid style with no complex formatting
                        try {
                            $table.Style = "Table Grid"
                        } 
                        catch {
                            Write-Host "Could not apply table style"
                        }
                        
                        # 5. Set up the header row with only the subset of headers for this table
                        for ($i = 0; $i -lt $currentColCount; $i++) {
                            $headerIdx = $startCol + $i
                            if ($headerIdx -lt $headers.Count) {
                                $headerText = $headers[$headerIdx]
                                $table.Cell(1, $i + 1).Range.Text = $headerText
                                $table.Cell(1, $i + 1).Range.Bold = $true
                            }
                        }
                        
                        # Add light gray shading to header row
                        try {
                            $table.Rows.Item(1).Shading.BackgroundPatternColor = 14277081 # Very light gray
                        } 
                        catch {
                            Write-Host "Could not set header row shading"
                        }
                        
                        # 6. Add data to the table - only the columns for this sub-table
                        for ($row = 0; $row -lt $dataRows.Count; $row++) {
                            $cells = $dataRows[$row]
                            
                            # Fill each cell with data for this column subset
                            for ($i = 0; $i -lt $currentColCount; $i++) {
                                $colIdx = $startCol + $i
                                if ($colIdx -lt $cells.Count) {
                                    $cellText = $cells[$colIdx].Trim()
                                    # Clean up markdown formatting for cell text
                                    $cellText = $cellText.Replace("**", "").Replace("*", "").Replace("__", "").Replace("_", "")
                                    $table.Cell($row + 2, $i + 1).Range.Text = $cellText
                                }
                            }
                        }
                        
                        # 7. Set table formatting - basic but reliable
                        try {
                            # Ensure word wrap is enabled in cells
                            foreach ($cell in $table.Range.Cells) {
                                $cell.WordWrap = $true
                            }
                            
                            # Auto-fit to content for better display
                            $table.AutoFitBehavior(1) # wdAutoFitContent
                            
                            # Allow rows to break across pages
                            $table.Rows.AllowBreakAcrossPages = $true
                        }
                        catch {
                            Write-Host "Error in table formatting: $_"
                        }
                    }
                    catch {
                        Write-Host "Error creating table $($tableIndex + 1): $_"
                        
                        # Fallback - output as formatted text for this section
                        if ($tableIndex == 0) { # Only output as text for the first table
                            $Selection.Font.Bold = $true
                            foreach ($header in $headers) {
                                $Selection.TypeText($header + "  ")
                            }
                            $Selection.Font.Bold = $false
                            $Selection.TypeParagraph()
                            
                            foreach ($dataRow in $dataRows) {
                                foreach ($cell in $dataRow) {
                                    $Selection.TypeText($cell + "  ")
                                }
                                $Selection.TypeParagraph()
                            }
                        }
                    }
                }
                
                # Add space after all tables
                $Selection.TypeParagraph()
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
    
    # Set up document for optimal readability
    $doc.PageSetup.LeftMargin = 36 # 0.5 inch in points
    $doc.PageSetup.RightMargin = 36
    $doc.PageSetup.TopMargin = 72
    $doc.PageSetup.BottomMargin = 72
    
    # Use landscape orientation for wider tables
    $doc.PageSetup.Orientation = 1 # wdOrientLandscape
    
    # Set a clean, readable font
    $word.Selection.Font.Name = "Calibri"
    $word.Selection.Font.Size = 11
    
    # Ensure white background for good contrast
    try {
        $doc.Background.Fill.Visible = $true
        $doc.Background.Fill.ForeColor.RGB = 16777215 # White
    } catch {
        Write-Host "Could not set document background"
    }
    
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
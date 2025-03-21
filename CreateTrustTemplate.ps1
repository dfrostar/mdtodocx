# Simple script to create a trust inventory template with empty tables
param (
    [Parameter(Mandatory = $false)]
    [string]$InputFile = "trust-asset-inventory.txt",
    
    [Parameter(Mandatory = $false)]
    [string]$OutputFile = "trust-template.docx"
)

# Function to check if a line is a section heading
function Is-SectionHeading {
    param ([string]$Line)
    
    # Skip lines that contain pipe characters - these are likely table rows, not headings
    if ($Line -match '\|') {
        return $false
    }
    
    # Remove markdown symbols before checking
    $cleanLine = $Line -replace '^#+\s*', ''
    
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
        if ($cleanLine -match $pattern) {
            return $true
        }
    }
    
    return $false
}

# Function to clean a section heading (remove markdown symbols)
function Clean-SectionHeading {
    param ([string]$Line)
    
    return $Line -replace '^#+\s*', ''
}

# Function to check if a line contains table headers (has pipe characters)
function Has-TableHeaders {
    param ([string]$Line)
    
    # Lines with multiple pipe characters are potential table headers
    if ($Line -match '\|.*\|') {
        $pipeCount = ($Line.ToCharArray() | Where-Object { $_ -eq '|' } | Measure-Object).Count
        if ($pipeCount -ge 2) {
            return $true
        }
    }
    
    return $false
}

# Function to extract headers from a pipe-delimited line
function Extract-Headers {
    param ([string]$Line)
    
    $line = $Line.Trim()
    
    # Remove starting and ending pipes
    if ($line.StartsWith('|')) {
        $line = $line.Substring(1)
    }
    if ($line.EndsWith('|')) {
        $line = $line.Substring(0, $line.Length - 1)
    }
    
    # Split by pipe and clean up
    $headers = $line.Split('|') | ForEach-Object { $_.Trim() }
    
    # Remove any markdown formatting and filter out empty headers
    $cleanHeaders = @()
    foreach ($header in $headers) {
        if (-not [string]::IsNullOrWhiteSpace($header)) {
            $cleanHeader = $header.Replace("**", "").Replace("*", "").Replace("__", "").Replace("_", "")
            $cleanHeaders += $cleanHeader
        }
    }
    
    return $cleanHeaders
}

# Function to ensure we're not in a table
function Ensure-NotInTable {
    param ($Word)
    
    # Check if we're in a table
    if ($Word.Selection.Information(12)) { # 12 = wdWithInTable
        Write-Host "Currently in a table - moving out..."
        
        # Try multiple methods to ensure we're completely out of the table
        $Word.Selection.EndKey(6, 0) # 6 = wdStory - moves to end of document
        $Word.Selection.TypeParagraph()
        $Word.Selection.TypeParagraph()
        
        # Double-check we're not in a table
        if ($Word.Selection.Information(12)) {
            Write-Host "Still in a table after first attempt - trying alternative method..."
            $Word.Selection.MoveRight(1, 1, 1) # Move outside the table
            $Word.Selection.TypeParagraph()
            $Word.Selection.TypeParagraph()
        }
        
        # Final check
        if ($Word.Selection.Information(12)) {
            Write-Host "WARNING: Still unable to exit table. Using more aggressive method..."
            # Create a new paragraph after the table
            $range = $Word.Selection.Range
            $range.MoveEnd(1, 1) # 1 = wdCharacter
            $range.Collapse(0) # 0 = wdCollapseEnd
            $range.InsertParagraphAfter()
            $range.InsertParagraphAfter()
            $range.Select()
        }
    }
}

# Main script
try {
    Write-Host "Starting trust template creation..."
    
    # Check if input file exists
    if (-not (Test-Path $InputFile)) {
        throw "Input file not found: $InputFile"
    }
    
    # Read the file content
    $lines = Get-Content -Path $InputFile
    Write-Host "Read $($lines.Count) lines from input file"
    
    # Initialize Word
    Write-Host "Initializing Word..."
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    
    # Create a new document
    $doc = $word.Documents.Add()
    
    # Set document properties
    $doc.PageSetup.Orientation = 1 # wdOrientLandscape
    $doc.PageSetup.LeftMargin = 36 # 0.5 inch
    $doc.PageSetup.RightMargin = 36
    $doc.PageSetup.TopMargin = 36 # 0.5 inch
    $doc.PageSetup.BottomMargin = 36 # 0.5 inch
    
    # Set default font
    $word.Selection.Font.Name = "Calibri"
    $word.Selection.Font.Size = 11
    
    # Add title
    $word.Selection.Font.Size = 18
    $word.Selection.Font.Bold = $true
    $word.Selection.TypeText("Trust Asset Inventory Template")
    $word.Selection.TypeParagraph()
    $word.Selection.TypeParagraph()
    $word.Selection.Font.Bold = $false
    $word.Selection.Font.Size = 11
    
    # Define font sizes
    $sectionHeadingSize = 14 # Section headings
    $tableHeaderSize = 12 # Table column headers
    
    # Track sections we've already processed to avoid duplicates
    $processedSections = @{}
    
    # Process the file to find sections and table headers
    $currentSection = ""
    $table = $null
    
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        
        # Skip empty lines and lines with just pipes and spaces
        if ([string]::IsNullOrWhiteSpace($line) -or $line -match '^\s*\|\s*\|\s*$') {
            continue
        }
        
        # Check if this is a section heading
        if (Is-SectionHeading $line) {
            $currentSection = Clean-SectionHeading $line
            
            # Skip if we've already processed this section
            if ($processedSections.ContainsKey($currentSection)) {
                Write-Host "Skipping duplicate section: $currentSection"
                continue
            }
            
            # Mark this section as processed
            $processedSections[$currentSection] = $true
            
            Write-Host "Found section: $currentSection"
            
            # Ensure we're not in a table before adding a section heading
            Ensure-NotInTable -Word $word
            
            # Add section heading with proper formatting
            $word.Selection.Font.Size = $sectionHeadingSize
            $word.Selection.Font.Bold = $true
            $word.Selection.Font.Color = 5460735 # Dark blue
            $word.Selection.TypeText($currentSection)
            $word.Selection.TypeParagraph()
            $word.Selection.Font.Bold = $false
            $word.Selection.Font.Color = 0 # Black
            $word.Selection.Font.Size = 11
            
            # Look for table headers in the next few lines
            $foundTable = $false
            for ($j = $i + 1; $j -lt [Math]::Min($i + 5, $lines.Count); $j++) {
                if (Has-TableHeaders $lines[$j]) {
                    $headerLine = $lines[$j]
                    $headers = Extract-Headers $headerLine
                    
                    if ($headers.Count -gt 0) {
                        $foundTable = $true
                        Write-Host "Found table headers: $($headers -join ', ')"
                        
                        # Create a table with these headers and 7 empty rows
                        $rowCount = 8 # Header + 7 empty rows
                        $colCount = $headers.Count
                        
                        # Limit to 10 columns max for safety
                        if ($colCount -gt 10) {
                            Write-Host "Limiting to 10 columns (was $colCount)"
                            $colCount = 10
                            $headers = $headers[0..9]
                        }
                        
                        Write-Host "Creating table with $rowCount rows and $colCount columns"
                        
                        try {
                            # Create the table
                            $table = $word.Selection.Tables.Add($word.Selection.Range, $rowCount, $colCount)
                            $table.Borders.Enable = $true
                            
                            # Set table properties for better display
                            $table.Borders.InsideLineStyle = 1 # wdLineStyleSingle
                            $table.Borders.OutsideLineStyle = 1 # wdLineStyleSingle
                            
                            # Add headers to the first row
                            for ($col = 0; $col -lt $headers.Count -and $col -lt $colCount; $col++) {
                                $headerText = [string]$headers[$col]  # Ensure it's a string
                                $table.Cell(1, $col + 1).Range.Text = $headerText
                                
                                # Format header cell
                                $table.Cell(1, $col + 1).Range.Font.Bold = $true
                                $table.Cell(1, $col + 1).Range.Font.Size = $tableHeaderSize
                            }
                            
                            # Format header row
                            $table.Rows.Item(1).Shading.BackgroundPatternColor = 14277081 # Light gray
                            
                            # Auto-fit table to content
                            $table.AutoFitBehavior(1) # wdAutoFitContent
                            
                            # Ensure cell word-wrap is enabled
                            foreach ($cell in $table.Range.Cells) {
                                $cell.WordWrap = $true
                            }
                            
                            # This is critical - we need to move the cursor AFTER the table to avoid nesting
                            $range = $table.Range
                            $range.Collapse(0) # 0 = wdCollapseEnd
                            $range.Select()
                            $word.Selection.TypeParagraph()
                            $word.Selection.TypeParagraph()
                        }
                        catch {
                            Write-Host "Error creating table for $currentSection"
                            # Just add a paragraph and continue
                            $word.Selection.TypeParagraph()
                        }
                        
                        # Skip ahead in the file
                        $i = $j
                        break
                    }
                }
            }
            
            # If no table was found, ensure we still have proper spacing
            if (-not $foundTable) {
                $word.Selection.TypeParagraph()
            }
        }
    }
    
    # Specifically add any important sections that might be missing
    $coreSections = @(
        "Professional Contact Information",
        "Important Dates & Deadlines"
    )
    
    foreach ($section in $coreSections) {
        if (-not $processedSections.ContainsKey($section)) {
            Write-Host "Adding missing core section: $section"
            
            # Ensure we're not in a table
            Ensure-NotInTable -Word $word
            
            # Add section heading
            $word.Selection.Font.Size = $sectionHeadingSize
            $word.Selection.Font.Bold = $true
            $word.Selection.Font.Color = 5460735 # Dark blue
            $word.Selection.TypeText($section)
            $word.Selection.TypeParagraph()
            $word.Selection.Font.Bold = $false
            $word.Selection.Font.Color = 0 # Black
            $word.Selection.Font.Size = 11
            
            # Add an appropriate table based on the section
            $headers = @()
            
            if ($section -eq "Professional Contact Information") {
                $headers = @("Role", "Name", "Company", "Phone", "Email", "Notes")
            }
            elseif ($section -eq "Important Dates & Deadlines") {
                $headers = @("Description", "Due Date", "Responsible Party", "Status", "Notes")
            }
            
            if ($headers.Count -gt 0) {
                # Create the table
                $rowCount = 8 # Header + 7 empty rows
                $colCount = $headers.Count
                
                try {
                    $table = $word.Selection.Tables.Add($word.Selection.Range, $rowCount, $colCount)
                    $table.Borders.Enable = $true
                    
                    # Set table properties
                    $table.Borders.InsideLineStyle = 1
                    $table.Borders.OutsideLineStyle = 1
                    
                    # Add headers
                    for ($col = 0; $col -lt $headers.Count; $col++) {
                        $table.Cell(1, $col + 1).Range.Text = $headers[$col]
                        $table.Cell(1, $col + 1).Range.Font.Bold = $true
                        $table.Cell(1, $col + 1).Range.Font.Size = $tableHeaderSize
                    }
                    
                    # Format header row
                    $table.Rows.Item(1).Shading.BackgroundPatternColor = 14277081 # Light gray
                    
                    # Auto-fit and enable word-wrap
                    $table.AutoFitBehavior(1)
                    foreach ($cell in $table.Range.Cells) {
                        $cell.WordWrap = $true
                    }
                    
                    # Move past the table
                    $range = $table.Range
                    $range.Collapse(0) # 0 = wdCollapseEnd
                    $range.Select()
                    $word.Selection.TypeParagraph()
                    $word.Selection.TypeParagraph()
                }
                catch {
                    Write-Host "Error creating table for $section"
                    $word.Selection.TypeParagraph()
                }
            }
            
            # Mark as processed
            $processedSections[$section] = $true
        }
    }
    
    # Save and close
    Write-Host "Saving document to: $OutputFile"
    $absolutePath = [System.IO.Path]::GetFullPath($OutputFile)
    $doc.SaveAs([ref]$absolutePath)
    $doc.Close()
    $word.Quit()
    
    # Release COM objects
    if ($table -ne $null) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($table) | Out-Null } catch { }
    }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    Write-Host "Trust template created successfully: $OutputFile" -ForegroundColor Green
}
catch {
    Write-Host "ERROR: $_" -ForegroundColor Red
    
    # Clean up
    if ($doc -ne $null) { try { $doc.Close($false) } catch { } }
    if ($word -ne $null) { try { $word.Quit() } catch { } }
    
    throw
} 
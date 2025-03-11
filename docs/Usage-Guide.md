# Usage Guide

This guide explains how to use the md_txt-to_docx tools for converting text files to formatted Word documents.

## Available Conversion Tools

This repository includes multiple tools for converting text files to Word documents:

### 1. Convert-TextToDocx.ps1 (formerly Convert-Trust.ps1)

The main general-purpose converter that can handle various text formats. It currently leverages the trust inventory converter but will be generalized in future versions.

```powershell
.\Convert-TextToDocx.ps1 -InputFile "your-file.txt" -OutputFile "output.docx" [-ShowWord]
```

Parameters:
- `-InputFile`: (Required) Path to the input text file
- `-OutputFile`: (Optional) Path for the output Word document. If not specified, it uses the input filename with a .docx extension
- `-ShowWord`: (Optional) Shows the Word application during conversion (useful for debugging)

### 2. Test-TrustInventory.ps1

A specialized converter optimized for trust inventory documents but designed to be extensible to other formats.

```powershell
.\Test-TrustInventory.ps1 -InputFile "trust-asset-inventory.txt" -OutputFile "output.docx" [-ShowWord]
```

Parameters:
- `-InputFile`: (Optional) Path to the input text file (defaults to "trust-asset-inventory.txt")
- `-OutputFile`: (Optional) Path for the output Word document (defaults to "trust-asset-inventory-test.docx")
- `-ShowWord`: (Optional) Shows the Word application during conversion

### 3. CreateTrustTemplate.ps1

Creates an empty template document with proper formatting for trust inventory data.

```powershell
.\CreateTrustTemplate.ps1 -InputFile "trust-asset-inventory.txt" -OutputFile "template.docx"
```

Parameters:
- `-InputFile`: (Optional) Path to the reference text file (defaults to "trust-asset-inventory.txt")
- `-OutputFile`: (Optional) Path for the output template (defaults to "trust-template.docx")

## Input File Format

The converter works best with Markdown-formatted text files or plain text files with a consistent structure:

### Headings

Use standard Markdown heading format:

```markdown
# Level 1 Heading
## Level 2 Heading
### Level 3 Heading
```

### Tables

Tables should use the pipe character (|) to separate columns:

```markdown
| Column 1 | Column 2 | Column 3 |
|----------|----------|----------|
| Data 1   | Data 2   | Data 3   |
| More data| More info| Details  |
```

The converter will automatically detect tables and format them appropriately in the output document.

## Examples

### Converting a Trust Inventory Document

```powershell
.\Convert-TextToDocx.ps1 -InputFile "trust-asset-inventory.txt" -OutputFile "trust-asset-inventory.docx"
```

### Creating a Trust Template

```powershell
.\CreateTrustTemplate.ps1 -OutputFile "my-template.docx"
```

### Converting with Visible Word Application (for debugging)

```powershell
.\Convert-TextToDocx.ps1 -InputFile "document.txt" -OutputFile "document.docx" -ShowWord
```

## Advanced Features

### Document Structure Analysis

The converters perform a two-pass analysis:
1. First, they analyze the structure of the document to identify headings, sections, tables, and other elements
2. Then, they create a Word document with proper formatting based on the analysis

This approach ensures that the resulting Word document maintains the logical structure of the original text file.

### Table Detection

The converters use intelligent table detection to:
- Find tables in the input text
- Determine table structure, including headers
- Format tables appropriately in Word
- Skip separator lines and placeholder content

### Section Handling

Special attention is given to section headings to ensure they're not embedded within tables and are properly formatted with:
- Larger font size
- Bold formatting
- Color coding
- Proper spacing

## Troubleshooting

### Common Issues

1. **Tables embedded in columns**: If you see section headers appearing inside table cells, your input file might not have proper spacing between sections. Add empty lines between sections and tables.

2. **Word application not closing**: If the Word application remains open after conversion, you may need to manually close it. This can happen if the conversion process encounters an error.

3. **Conversion errors**: Check that your input file follows the expected format, especially for tables. Inconsistent formatting can cause issues.

### Getting Help

If you encounter problems:
1. Check if your issue is already reported in the GitHub Issues section
2. Create a new Issue with detailed information about your problem
3. Include error messages and the input file that caused the issue

## Future Development

The converters are being actively developed to support:
- More general-purpose text formats
- Enhanced table detection and formatting
- Plugin architecture for custom document types
- Additional output formatting options

Your feedback and contributions are welcome!


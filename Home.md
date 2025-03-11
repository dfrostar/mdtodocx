# Welcome to the md_txt-to_docx Wiki

This tool allows you to convert Markdown (.md) and plain text (.txt) files to properly formatted Microsoft Word (.docx) documents, with special focus on handling tables properly.

## Quick Navigation

- [Installation](./Installation)
- [Usage Guide](./Usage-Guide)
- [File Format Support](./File-Format-Support)
- [Examples](./Examples)
- [Troubleshooting](./Troubleshooting)
- [Advanced Usage](./Advanced-Usage)
- [Contributing](./Contributing)

## Overview

The md_txt-to_docx tool is a PowerShell-based converter that transforms Markdown and text files into well-formatted Word documents. It preserves formatting, handles tables, lists, and basic text styling.

## Key Features

- **Multiple File Format Support**: Converts both Markdown (.md) and plain text (.txt) files
- **Enhanced Table Support**: Correctly formats and preserves tables in Word using pipe-delimited formats
- **Heading Styles**: Applies appropriate Word heading styles
- **Text Preservation**: Maintains the structure of your documents
- **Easy to Use**: Simple PowerShell script with batch file wrappers
- **Verbose Logging Option**: Detailed logging for troubleshooting

## Quick Start

1. Clone the repository:
   ```
   git clone https://github.com/dfrostar/md_txt-to_docx.git
   ```

2. Run the script with your Markdown or TXT file:
   - Using the main script (recommended for most users):
   ```
   .\Convert-MarkdownToWord.bat path\to\your\file.txt
   ```
   
   - For more control with the simplified table converter:
   ```
   .\Convert-TextToWord.bat path\to\your\file.txt output.docx
   ```
   
   - For direct PowerShell access with verbose logging:
   ```powershell
   .\Convert-TxtToWord.ps1 -InputFile "path\to\your\file.txt" -OutputFile "output.docx" -Verbose
   ```

## Example

Converting a text file with a pipe-delimited table to Word:

```
| Column 1       | Column 2       | Column 3      |
|----------------|----------------|---------------|
| Row 1, Cell 1  | Row 1, Cell 2  | Row 1, Cell 3 |
| Row 2, Cell 1  | Row 2, Cell 2  | Row 2, Cell 3 |
```

Using the command:
```
.\Convert-MarkdownToWord.bat inventory.txt inventory.docx
```

Results in a properly formatted Word document with a styled table.

## Latest Updates

- **New SimpleTableConverter.ps1**: A streamlined script focused on correct table processing
- **Improved Pipe-Delimited Table Handling**: Better detection and formatting of tables
- **New Convert-TextToWord.bat**: Additional batch wrapper for the simplified converter
- **Enhanced Error Logging**: Better troubleshooting with the verbose logging option
- **Updated README**: Comprehensive documentation and examples

## Requirements

- Windows operating system
- PowerShell 5.1 or higher
- Microsoft Word installed 
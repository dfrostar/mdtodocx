# Welcome to the mdtodocx Wiki

This tool allows you to convert Markdown (.md) and plain text (.txt) files to properly formatted Microsoft Word (.docx) documents.

## Quick Navigation

- [Installation](./Installation)
- [Usage Guide](./Usage-Guide)
- [File Format Support](./File-Format-Support)
- [Examples](./Examples)
- [Troubleshooting](./Troubleshooting)
- [Advanced Usage](./Advanced-Usage)
- [Contributing](./Contributing)

## Overview

The mdtodocx tool is a PowerShell-based converter that transforms Markdown and text files into well-formatted Word documents. It preserves formatting, handles tables, lists, and basic text styling.

## Key Features

- **Multiple File Format Support**: Converts both Markdown (.md) and plain text (.txt) files
- **Table Support**: Correctly formats and preserves tables in Word
- **Heading Styles**: Applies appropriate Word heading styles
- **List Formatting**: Preserves both ordered and unordered lists
- **Text Styling**: Maintains bold, italic, and other basic text formatting
- **Easy to Use**: Simple PowerShell script with a batch file wrapper

## Quick Start

1. Clone the repository:
   ```
   git clone https://github.com/dfrostar/mdtodocx.git
   ```

2. Run the script with your Markdown or TXT file:
   ```powershell
   .\Convert-MarkdownToWord.ps1 -InputFile "path\to\your\file.md" -OutputFile "path\to\output.docx"
   ```
   
   Or use the batch file:
   ```
   .\Convert-MarkdownToWord.bat path\to\your\file.txt
   ```

## Example

Converting a text file with a table to Word:

```
+----------------+----------------+---------------+
| Column 1       | Column 2       | Column 3      |
+----------------+----------------+---------------+
| Row 1, Cell 1  | Row 1, Cell 2  | Row 1, Cell 3 |
| Row 2, Cell 1  | Row 2, Cell 2  | Row 2, Cell 3 |
+----------------+----------------+---------------+
```

Using the command:
```
.\Convert-MarkdownToWord.bat inventory.txt inventory.docx
```

Results in a properly formatted Word document with a styled table.

## Latest Updates

- Added support for TXT files with tables
- Improved table formatting
- Updated batch file wrapper to support both file types

## Requirements

- Windows operating system
- PowerShell 5.1 or higher
- Microsoft Word installed 
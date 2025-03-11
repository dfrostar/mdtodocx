# Markdown and Text to Word Converter

A PowerShell script for converting Markdown and plain text files to properly formatted Microsoft Word documents, preserving tables, lists, and text styling.

## Overview

This converter is a simple yet powerful tool designed to transform Markdown and plain text documents into properly formatted Microsoft Word documents. It handles tables, headings, lists, and basic formatting, making it ideal for generating professional documents from simple Markdown or TXT files.

## Features

- **Multiple File Format Support**: Converts both Markdown (.md) and plain text (.txt) files
- **Table Support**: Correctly formats and preserves tables in Word from both formats
- **Heading Styles**: Applies appropriate Word heading styles to headings
- **List Formatting**: Preserves both ordered and unordered lists
- **Text Styling**: Maintains bold, italic, and other basic text formatting
- **Parameterized Script**: Easy to use with customizable input and output paths

## Requirements

- Windows operating system
- PowerShell 5.1 or higher
- Microsoft Word installed

## Quick Start

1. Clone this repository:
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

## Examples

Convert a Markdown file to Word:
```powershell
.\Convert-MarkdownToWord.ps1 -InputFile "C:\Users\username\Documents\document.md" -OutputFile "C:\Users\username\Documents\document.docx"
```

Convert a plain text file with tables to Word:
```powershell
.\Convert-MarkdownToWord.ps1 -InputFile "C:\Users\username\Documents\inventory.txt" -OutputFile "C:\Users\username\Documents\inventory.docx"
```

Show the Word application during conversion (useful for debugging):
```powershell
.\Convert-MarkdownToWord.ps1 -InputFile "document.md" -OutputFile "document.docx" -ShowWord
```

Using the batch file wrapper:
```
Convert-MarkdownToWord.bat inventory.txt
Convert-MarkdownToWord.bat document.md output.docx show
```

## Documentation

For detailed usage instructions, see the [Wiki](https://github.com/dfrostar/mdtodocx/wiki).

## Limitations

- Complex nested structures may not convert perfectly
- Limited support for advanced Markdown features like code blocks
- Images are not currently supported

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Thanks to the PowerShell and Word automation communities for their documentation and examples.
- Inspired by the need for better document conversion tools with proper table support. 

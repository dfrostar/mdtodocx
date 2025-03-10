# Markdown to Word Converter

A PowerShell script for converting Markdown files to properly formatted Microsoft Word documents, preserving tables, lists, and text styling.

## Overview

The Markdown to Word Converter is a simple yet powerful tool designed to transform Markdown documents into properly formatted Microsoft Word documents. It handles tables, headings, lists, and basic formatting, making it ideal for generating professional documents from simple Markdown files.

## Features

- **Table Support**: Correctly formats and preserves Markdown tables in Word
- **Heading Styles**: Applies appropriate Word heading styles to Markdown headings
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
   git clone https://github.com/mdtodocx/markdown-to-word.git
   ```

2. Run the script with your Markdown file:
   ```powershell
   .\Convert-MarkdownToWord.ps1 -InputFile "path\to\your\markdown.md" -OutputFile "path\to\output.docx"
   ```

## Examples

Convert a Markdown file to Word:
```powershell
.\Convert-MarkdownToWord.ps1 -InputFile "C:\Users\username\Documents\document.md" -OutputFile "C:\Users\username\Documents\document.docx"
```

Show the Word application during conversion (useful for debugging):
```powershell
.\Convert-MarkdownToWord.ps1 -InputFile "document.md" -OutputFile "document.docx" -ShowWord
```

## Documentation

For detailed usage instructions, see the [USAGE.md](USAGE.md) file.

## Limitations

- Complex nested Markdown structures may not convert perfectly
- Limited support for advanced Markdown features like code blocks
- Images are not currently supported

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Thanks to the PowerShell and Word automation communities for their documentation and examples.
- Inspired by the need for better Markdown to Word conversion tools with proper table support. 

# Markdown & TXT to Word Document Converter

A set of PowerShell tools for converting Markdown and plain text files into professionally formatted Microsoft Word documents with proper table formatting.

## Features

- **Semantic understanding** of document structure
- **Intelligent table detection** and formatting
- **Special handling for trust inventory documents**
- **Command-line interface** via batch files
- **Detailed logging** for troubleshooting

## Tools Included

### 1. SemanticDocConverter.ps1
General-purpose converter that analyzes the semantic structure of documents before converting them to Word format.

### 2. TrustDocConverter.ps1
Specialized converter for trust inventory documents with optimized handling of headings, sections, and tables common in trust inventories.

### 3. Convert-Semantic.bat
Batch wrapper for the semantic converter that simplifies command-line usage.

### 4. Convert-TrustInventory.bat
Batch wrapper specifically for trust inventory documents with specialized formatting.

## Usage

### Converting a standard Markdown file:

```batch
Convert-Semantic.bat input.md output.docx [debug]
```

### Converting a text file with general structure:

```batch
Convert-Semantic.bat input.txt output.docx [debug]
```

### Converting a trust inventory document:

```batch
Convert-TrustInventory.bat trust-inventory.txt [output.docx] [debug]
```

## Benefits Over Basic Conversion

- **Two-pass processing**: Analyzes document structure first, then formats it
- **True document understanding**: Identifies real tables vs. section headings
- **Contextual processing**: Understands what different document elements mean
- **Proper table detection**: Only treats actual tabular data as tables
- **Smart content handling**: Skips empty rows and separator lines

## Requirements

- Windows operating system
- PowerShell 5.1 or later
- Microsoft Word installed

## Support

For issues or feature requests, please submit an issue on the GitHub repository.

## License

MIT License

## Files

- `SimpleTableConverter.ps1` - The main PowerShell script for conversion
- `Convert-TextToWord.bat` - Batch file wrapper for easier command-line use

## Input File Format

The input text file should use:

- Headings with `#` symbols (e.g., `# Heading 1`, `## Heading 2`)
- Tables in pipe-delimited format:

```
| Header 1 | Header 2 | Header 3 |
|----------|----------|----------|
| Data 1   | Data 2   | Data 3   |
```

## Example

For the file `trust-asset-inventory.txt`:

```
# Trust Asset Inventory

## Financial Accounts
| Account Type | Institution | Account Number |
|--------------|-------------|----------------|
| Checking     | First Bank  | 1234           |
| Savings      | Credit Union| 5678           |
```

Convert it with:

```
Convert-TextToWord.bat trust-asset-inventory.txt trust-asset-inventory.docx
```

The resulting Word document will have a properly formatted heading hierarchy and tables with appropriate styling.

## Troubleshooting

If you encounter issues:

1. Ensure Microsoft Word is properly installed
2. Check that your input file uses correct pipe-delimited table formatting
3. Try running the PowerShell script directly to see detailed error messages

## Acknowledgments

- Thanks to the PowerShell and Word automation communities for their documentation and examples.
- Inspired by the need for better document conversion tools with proper table support. 

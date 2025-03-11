# Text to Word Document Converter

This utility converts text files with tables (pipe-delimited format) to Microsoft Word documents, preserving formatting, headings, and table structures.

## Files

- `SimpleTableConverter.ps1` - The main PowerShell script for conversion
- `Convert-TextToWord.bat` - Batch file wrapper for easier command-line use

## Features

- Converts Markdown-style headings (# Heading 1, ## Heading 2, etc.)
- Properly formats pipe-delimited tables with header rows
- Preserves plain text paragraphs
- Auto-fits table columns to content
- Applies proper styling (bold headers, table borders)

## Requirements

- Windows OS
- Microsoft Word installed
- PowerShell 5.1 or higher

## Usage

### PowerShell Method

```powershell
powershell -ExecutionPolicy Bypass -File SimpleTableConverter.ps1 -InputFile "your-input.txt" -OutputFile "your-output.docx"
```

Optional parameters:
- `-ShowWord`: Shows the Word application during conversion

### Batch File Method

```
Convert-TextToWord.bat input.txt output.docx
```

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

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Thanks to the PowerShell and Word automation communities for their documentation and examples.
- Inspired by the need for better document conversion tools with proper table support. 

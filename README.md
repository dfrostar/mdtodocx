# Markdown & TXT to Word Document Converter

A set of PowerShell tools for converting Markdown and plain text files into professionally formatted Microsoft Word documents with proper table formatting.

## Features

- **Semantic understanding** of document structure
- **Intelligent table detection** and formatting
- **Special handling for trust inventory documents**
- **Command-line interface** via PowerShell scripts
- **Detailed logging** for troubleshooting

## Tools Included

### 1. Convert-TextToDocx.ps1
General-purpose converter that can handle various text formats. Currently uses the trust inventory converter but designed to be extensible to other formats.

### 2. Test-TrustInventory.ps1
Specialized converter for trust inventory documents with optimized handling of headings, sections, and tables common in trust inventories.

### 3. CreateTrustTemplate.ps1
Creates empty template documents with proper structure for trust inventory data.

## Usage

### Converting a standard text file:

```powershell
.\Convert-TextToDocx.ps1 -InputFile input.txt -OutputFile output.docx
```

### Converting with visible Word application (for debugging):

```powershell
.\Convert-TextToDocx.ps1 -InputFile input.txt -OutputFile output.docx -ShowWord
```

### Creating a trust inventory template:

```powershell
.\CreateTrustTemplate.ps1 -OutputFile template.docx
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

## Documentation

Detailed documentation is available in the `docs` directory:

- [Usage Guide](docs/Usage-Guide.md) - Comprehensive guide to using the converters
- [Architecture](docs/Architecture.md) - Details about the project's internal architecture

## Support

For issues or feature requests, please submit an issue on the GitHub repository.

## License

MIT License

## Files

- `Convert-TextToDocx.ps1` - Main entry point for general text conversion
- `Test-TrustInventory.ps1` - Specialized converter for trust inventory documents
- `CreateTrustTemplate.ps1` - Template generator for trust documents

**PowerShell Module**:
- `DocConverter/` - PowerShell module with common functionality
- `DocConverter/Private/` - Internal utility functions
- `DocConverter/Public/` - Public module functions

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

### Real Estate Properties

| Property Address | Current Owner(s) | Estimated Value | Mortgage Balance | Mortgage Lender | Transfer Method | Date Transferred | Notes | Status |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 123 Main St, Anytown, USA | John & Jane Doe | $500,000 | $250,000 | First Bank | Quitclaim Deed | 3/15/2023 | Primary residence | Complete |
```

Convert it with:

```
.\Convert-TextToDocx.ps1 -InputFile trust-asset-inventory.txt
```

The resulting Word document will have properly formatted headings, sections, and tables with appropriate styling.

## Troubleshooting

If you encounter issues:

1. Ensure Microsoft Word is properly installed
2. Check that your input file uses correct pipe-delimited table formatting
3. Try running with the `-ShowWord` parameter to see what's happening during conversion
4. Check the console output for error messages

## Acknowledgments

- Thanks to the PowerShell and Word automation communities for their documentation and examples.
- Inspired by the need for better document conversion tools with proper table support.

## Future Development

- Enhanced plugin architecture
- Support for more document formats
- Improved table detection and formatting
- User interface improvements 

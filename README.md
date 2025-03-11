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

- `SemanticDocConverter.ps1` - General-purpose document converter with semantic understanding
- `TrustDocConverter.ps1` - Specialized converter for trust inventory documents
- `Convert-Semantic.bat` - Batch wrapper for the semantic converter
- `Convert-TrustInventory.bat` - Batch wrapper for the trust document converter
- `run-examples.bat` - Script to run all examples at once
- `Test-TrustInventory.ps1` - Standalone test script for trust inventory conversion

**PowerShell Module**:
- `DocConverter/` - PowerShell module with plugin architecture
- `Convert-Document.ps1` - Main script using the module
- `Convert-Document.bat` - Batch wrapper for the module script

## Input File Format

The input text file should use:

- Headings with `#` symbols (e.g., `# Heading 1`, `## Heading 2`)
- Tables in pipe-delimited format:

```
| Header 1 | Header 2 | Header 3 |
|----------|----------|----------|
| Data 1   | Data 2   | Data 3   |
```

## Advanced Features

### Plugin Architecture

The PowerShell module uses a plugin-based architecture that allows:
- Automatic detection of document types
- Specialized converters for different document formats
- Registering custom converters for new document types

### PowerShell Module Usage

```powershell
# Import the module
Import-Module .\DocConverter

# Auto-detect the converter type
Convert-DocumentToWord -InputFile "document.md" -OutputFile "document.docx"

# Specify a particular converter
Convert-DocumentToWord -InputFile "trust-inventory.txt" -ConverterName "TrustInventory" -ShowWord
```

### Batch File Usage

```batch
# Auto-detect converter type
Convert-Document.bat input.md output.docx

# Specify converter
Convert-Document.bat trust-inventory.txt output.docx TrustInventory

# Show Word during conversion
Convert-Document.bat document.md output.docx show
```

## Example

For the file `examples/sample-trust-inventory.txt`:

```
# Trust Asset Inventory

Real Estate Properties

| Property Address | Type | Value | Notes |
| 123 Main St, Anytown, USA | Primary Residence | $500,000 | Jointly owned with spouse |
```

Convert it with:

```
Convert-TrustInventory.bat examples\sample-trust-inventory.txt
```

The resulting Word document will have properly formatted headings, sections, and tables with appropriate styling.

## Examples

Check out the `examples` directory for sample input and output files:

- `sample-document.md` - General document with various formatting elements
- `sample-trust-inventory.txt` - Sample trust inventory document
- `run-examples.bat` - Batch script to convert all examples at once

## Troubleshooting

If you encounter issues:

1. Ensure Microsoft Word is properly installed
2. Check that your input file uses correct pipe-delimited table formatting
3. Try running the PowerShell script directly to see detailed error messages

## Acknowledgments

- Thanks to the PowerShell and Word automation communities for their documentation and examples.
- Inspired by the need for better document conversion tools with proper table support. 

# Usage Guide

This guide explains how to use the md_txt-to_docx converters to transform your text files into Word documents.

## Available Converters

The project includes three main ways to convert your files:

1. **Main Batch File** (`Convert-MarkdownToWord.bat`): The simplest option for most users
2. **Simple Table Converter** (`Convert-TextToWord.bat`): Optimized for table handling
3. **PowerShell Scripts** (`SimpleTableConverter.ps1` or `Convert-TxtToWord.ps1`): For advanced options

## Using the Main Batch File

This is the recommended option for most users:

```
.\Convert-MarkdownToWord.bat input.txt [output.docx]
```

### Parameters:
- `input.txt`: Your input text file (required)
- `output.docx`: The output Word document name (optional, defaults to same name as input with .docx extension)

### Example:
```
.\Convert-MarkdownToWord.bat trust-asset-inventory.txt
```

This will create `trust-asset-inventory.docx` in the same directory.

## Using the Simple Table Converter

If you're working primarily with tables, use this dedicated converter:

```
.\Convert-TextToWord.bat input.txt output.docx
```

### Parameters:
- `input.txt`: Your input text file (required)
- `output.docx`: The output Word document name (required)

### Example:
```
.\Convert-TextToWord.bat financial-data.txt financial-report.docx
```

## Using PowerShell Scripts Directly

For more control, you can use the PowerShell scripts directly:

### SimpleTableConverter.ps1 (Recommended)

```powershell
powershell -ExecutionPolicy Bypass -File .\SimpleTableConverter.ps1 -InputFile "input.txt" -OutputFile "output.docx" [-ShowWord]
```

### Parameters:
- `-InputFile`: Path to your input text file (required)
- `-OutputFile`: Path for the output Word document (required)
- `-ShowWord`: Optional switch to show Word during conversion (helpful for debugging)

### Example:
```powershell
powershell -ExecutionPolicy Bypass -File .\SimpleTableConverter.ps1 -InputFile "data.txt" -OutputFile "report.docx" -ShowWord
```

### Convert-TxtToWord.ps1 (With Detailed Logging)

```powershell
powershell -ExecutionPolicy Bypass -File .\Convert-TxtToWord.ps1 -InputFile "input.txt" -OutputFile "output.docx" -Verbose
```

### Parameters:
- `-InputFile`: Path to your input text file (required)
- `-OutputFile`: Path for the output Word document (required)
- `-ShowWord`: Optional switch to show Word during conversion
- `-Verbose`: Enables detailed logging to console and log file

### Example:
```powershell
powershell -ExecutionPolicy Bypass -File .\Convert-TxtToWord.ps1 -InputFile "complex-tables.txt" -OutputFile "report.docx" -Verbose
```

## Input File Requirements

Your input file should be:

1. A plain text file (.txt) or Markdown file (.md)
2. Properly formatted with:
   - Headings using # symbols (e.g., `# Heading 1`, `## Heading 2`)
   - Tables using pipe-delimited format (see [Table Formatting](./Table-Formatting))
   - Regular paragraphs with blank lines between them

## Output Options

The converter creates a standard Microsoft Word .docx file with:

- Proper heading styles applied
- Formatted tables with borders
- Preserved paragraph spacing
- Auto-fitted content

## Troubleshooting

If you encounter issues:

1. Try the `-Verbose` option with `Convert-TxtToWord.ps1` to get detailed logs
2. Check the `convert_log.txt` file for error messages
3. Verify that your input file is properly formatted
4. Make sure Microsoft Word is properly installed on your system

See the [Troubleshooting](./Troubleshooting) page for more details. 
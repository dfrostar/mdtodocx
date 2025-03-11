# Table Formatting Guide

This guide explains how to format tables in your text files for optimal conversion to Word documents using md_txt-to_docx.

## Supported Table Formats

The converter now supports two main types of table formats:

### 1. Pipe-Delimited Tables (Recommended)

This is the preferred format and works best with the improved table handling:

```
| Header 1 | Header 2 | Header 3 |
|----------|----------|----------|
| Data 1   | Data 2   | Data 3   |
| Data 4   | Data 5   | Data 6   |
```

### 2. Plus-Sign Border Tables (Legacy Support)

The original format with plus-sign borders is still supported:

```
+----------+----------+----------+
| Header 1 | Header 2 | Header 3 |
+----------+----------+----------+
| Data 1   | Data 2   | Data 3   |
| Data 4   | Data 5   | Data 6   |
+----------+----------+----------+
```

## Best Practices for Table Formatting

For the best results with the converter:

1. **Always include a header row** - The first row is treated as a header and will be formatted in bold
2. **Use separator rows** (like `|-----|-----|`) after headers for better readability in source files
3. **Keep consistent pipe characters** at the beginning and end of each row
4. **Ensure consistent column counts** across all rows
5. **Avoid merged cells** - These are not supported in the simple table format

## Example: Asset Inventory Table

Here's an example of a well-formatted table for an asset inventory:

```
| Asset Type | Description | Value | Owner | Location |
|------------|-------------|-------|-------|----------|
| Real Estate| Main House  |$350,000| John | 123 Main St |
| Vehicle    | 2022 Toyota |$28,000 | Jane | Garage |
| Financial  | 401(k)      |$125,000| John | Fidelity |
```

## Converting Tables

To convert files with tables:

### Using the Simplified Converter (Recommended)

```
.\Convert-TextToWord.bat inventory.txt inventory.docx
```

### Using the Original Converter with Verbose Logging

```
.\Convert-TxtToWord.ps1 -InputFile "inventory.txt" -OutputFile "inventory.docx" -Verbose
```

## Troubleshooting Table Issues

If your tables aren't converting correctly:

1. **Check the log file** - Look at `convert_log.txt` for detailed error messages
2. **Verify consistent column counts** - Make sure all rows have the same number of pipe characters
3. **Remove empty rows** - Blank rows in tables can sometimes cause issues
4. **Try the simplified converter** - `SimpleTableConverter.ps1` is optimized for table handling

## Limitations

- Tables must be surrounded by empty lines (before and after)
- Complex formatting within table cells may not be preserved
- Very wide tables might be truncated or reformatted to fit the page 
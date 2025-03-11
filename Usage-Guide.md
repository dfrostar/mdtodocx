# Document Converter Usage Guide

This guide explains how to use the different document converters in this repository for various use cases.

## 1. Simple Conversion (Using Batch Files)

### For General Documents

```batch
Convert-Semantic.bat input.md output.docx
```

Or for more detailed logging:

```batch
Convert-Semantic.bat input.md output.docx debug
```

### For Trust Inventory Documents

```batch
Convert-TrustInventory.bat trust-inventory.txt trust-inventory.docx
```

## 2. Advanced Usage (Direct PowerShell)

### General Document Converter

```powershell
# Basic usage
powershell -ExecutionPolicy Bypass -File SemanticDocConverter.ps1 -InputFile "document.md" -OutputFile "document.docx"

# With debug output and showing Word during conversion
powershell -ExecutionPolicy Bypass -File SemanticDocConverter.ps1 -InputFile "document.md" -OutputFile "document.docx" -Debug -ShowWord
```

### Trust Document Converter

```powershell
# Basic usage
powershell -ExecutionPolicy Bypass -File TrustDocConverter.ps1 -InputFile "trust-inventory.txt" -OutputFile "trust-inventory.docx"

# With verbose output and showing Word during conversion
powershell -ExecutionPolicy Bypass -File TrustDocConverter.ps1 -InputFile "trust-inventory.txt" -OutputFile "trust-inventory.docx" -Verbose -ShowWord
```

## 3. Input File Formats

### Markdown Format

The converter supports standard Markdown syntax:

```markdown
# Heading 1
## Heading 2
### Heading 3

Regular paragraph with **bold** and *italic* text.

| Header 1 | Header 2 | Header 3 |
|----------|----------|----------|
| Cell 1   | Cell 2   | Cell 3   |
| Cell 4   | Cell 5   | Cell 6   |
```

### Trust Inventory Format

For trust inventory documents, use a format like:

```
Real Estate Properties

| Property Address | Type | Value | Notes |
| 123 Main St, Anytown, USA | Primary Residence | $500,000 | Jointly owned with spouse |
| 456 Beach Ave, Coastville, USA | Vacation Home | $350,000 | Mortgage of $150,000 remaining |

Financial Accounts

| Institution | Account Type | Account Number | Value | Beneficiary |
| First National Bank | Checking | xxx-xxx-1234 | $25,000 | N/A |
| Investment Firm | Brokerage | yyy-yyyy-5678 | $250,000 | Spouse |
```

## 4. Troubleshooting

### Common Issues

1. **Table formatting issues**: 
   - Make sure your tables have consistent delimiters (pipes)
   - Check that all rows have the same number of columns

2. **Word automation errors**:
   - Ensure Microsoft Word is installed and functioning
   - Close any open Word documents before running the conversion

3. **Script execution policy**:
   - If you receive execution policy errors, run PowerShell as Administrator and execute:
     ```powershell
     Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
     ```

### Debug Output

For detailed processing information, add the `-Debug` parameter to SemanticDocConverter.ps1 or the `-Verbose` parameter to TrustDocConverter.ps1.

## 5. Output Customization

The converters apply standard formatting:
- Headings use Word's built-in heading styles
- Tables use the Table Grid style with borders
- Text uses the Normal style

To customize the output appearance, you can modify the style settings in the scripts.

## 6. Examples

See the `examples` directory for sample input files and their corresponding Word document outputs. 
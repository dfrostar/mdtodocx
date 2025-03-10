# Using the Markdown to Word Converter

This document provides detailed instructions on how to use the Markdown to Word converter script.

## Installation

1. Clone the repository or download the script:
   ```
   git clone https://github.com/yourusername/markdown-to-word.git
   ```

2. Navigate to the script directory:
   ```
   cd markdown-to-word
   ```

3. Ensure you have the necessary prerequisites:
   - Windows operating system
   - PowerShell 5.1 or higher
   - Microsoft Word installed

## Basic Usage

The script uses PowerShell parameters to specify the input Markdown file and output Word document.

```powershell
.\Convert-MarkdownToWord.ps1 -InputFile "path\to\your\markdown.md" -OutputFile "path\to\output.docx"
```

### Parameters

- `-InputFile`: The path to the Markdown file you want to convert (required).
- `-OutputFile`: The path where the Word document will be saved (required).
- `-ShowWord`: Optional switch to show the Word application during conversion (useful for debugging).

## Examples

### Convert a local Markdown file

```powershell
.\Convert-MarkdownToWord.ps1 -InputFile "C:\Documents\MyDocument.md" -OutputFile "C:\Documents\MyDocument.docx"
```

### Convert a file with visible Word application (for debugging)

```powershell
.\Convert-MarkdownToWord.ps1 -InputFile "C:\Documents\MyDocument.md" -OutputFile "C:\Documents\MyDocument.docx" -ShowWord
```

## Supported Markdown Features

The script supports the following Markdown features:

### Headings

Markdown headings are converted to Word heading styles:

```markdown
# Heading 1
## Heading 2
### Heading 3
```

### Tables

Markdown tables are correctly converted to Word tables:

```markdown
| Header 1 | Header 2 | Header 3 |
|----------|----------|----------|
| Cell 1   | Cell 2   | Cell 3   |
| Cell 4   | Cell 5   | Cell 6   |
```

### Lists

Both unordered and ordered lists are supported:

```markdown
- Item 1
- Item 2
- Item 3

1. First item
2. Second item
3. Third item
```

### Text Formatting

Basic text formatting is supported:

```markdown
**Bold text**
*Italic text*
```

## Limitations

- Complex nested Markdown formatting may not be perfectly preserved
- Images are not currently supported
- Some advanced Markdown features (code blocks, horizontal rules, etc.) may have limited support

## Troubleshooting

### Common Issues

1. **Word Not Installed**: Ensure that Microsoft Word is installed on your system.

2. **File Access Issues**: Make sure you have read access to the input file and write access to the output location.

3. **Error: "The file X is already open"**: Close any open instances of the output file in Word before running the script.

4. **Script Not Running**: Ensure PowerShell execution policy allows script execution:
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
   ```

### Getting Help

If you encounter issues, please:

1. Check the Issues section on GitHub to see if your problem has been reported
2. Open a new Issue with detailed information about your problem
3. Include error messages and steps to reproduce the issue

## Contributing

We welcome contributions! See the [CONTRIBUTING.md](CONTRIBUTING.md) file for details on how to contribute. 
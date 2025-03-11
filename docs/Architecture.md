# Project Architecture

This document explains the architecture of the md_txt-to_docx toolkit and how the different components interact.

## Overview

The md_txt-to_docx project is designed as a modular system for converting text files (both Markdown and plain text) into professionally formatted Microsoft Word documents. 

The architecture follows these key principles:
- **Separation of concerns**: Each script has a clear, focused responsibility
- **Two-pass conversion**: Analysis of document structure followed by formatting
- **Extensibility**: Common utilities and functions that can be reused
- **User-friendly interfaces**: Simple command-line tools with clear parameters

## Component Structure

The project consists of the following main components:

```
md_txt-to_docx/
├── Convert-TextToDocx.ps1       # Main entry point for general text conversion
├── Test-TrustInventory.ps1      # Specialized converter for trust documents
├── CreateTrustTemplate.ps1      # Template generator for trust documents
├── docs/                        # Documentation directory
│   ├── Usage-Guide.md           # Detailed usage instructions
│   ├── Architecture.md          # This document
│   └── USAGE.md                 # Legacy usage documentation
└── DocConverter/                # PowerShell module directory
    ├── Private/                 # Internal utility functions
    │   └── Common-Utils.ps1     # Shared utility functions
    └── Public/                  # Public module functions
```

## Process Flow

The general conversion process follows this flow:

1. **User invokes one of the conversion scripts** with input and output parameters
2. **Document analysis phase**:
   - Read the input file
   - Parse the document structure
   - Identify headings, sections, tables, and other elements
   - Create a structured representation of the document
3. **Document formatting phase**:
   - Initialize Word application
   - Create a new Word document
   - Format each element according to its type
   - Apply proper styling and layout
4. **Output and cleanup**:
   - Save the Word document
   - Clean up COM objects
   - Provide completion status

## Key Components

### Convert-TextToDocx.ps1

This is the main entry point for general text conversion. It:
- Processes command-line arguments
- Sets default output file if not specified
- Calls the appropriate converter (currently Test-TrustInventory.ps1)
- Handles errors and returns appropriate exit codes

### Test-TrustInventory.ps1

This script performs the actual conversion and contains:
- Document analysis functions
- Word document formatting functions
- Table detection and creation logic
- Section and heading formatting

### CreateTrustTemplate.ps1

This generates empty templates with proper formatting:
- Creates document structure from a reference file
- Generates empty tables with proper headers
- Formats section headings appropriately
- Ensures proper spacing between elements

### DocConverter Module

The PowerShell module provides a more structured approach:
- **Private utilities**: Common functions shared across converters
- **Public functions**: Functions that can be called by users

## Design Patterns

The project implements several design patterns:

### Two-Pass Processing

The converters use a two-pass approach:
1. **Analysis pass**: Read and understand document structure
2. **Formatting pass**: Create Word document based on analysis

This allows better understanding of the overall document before formatting begins.

### Function Specialization

Functions are specialized for specific tasks:
- `Is-SectionHeading`: Detects section headings
- `Is-TableData`: Identifies table data
- `Extract-TableHeaders`: Parses table headers
- `Analyze-TrustDocument`: Analyzes overall document structure
- `Format-TrustDocument`: Formats the document in Word

### Resource Management

Proper COM object management is implemented:
- `Initialize-WordApplication`: Creates Word instance
- `Close-WordObjects`: Properly releases COM objects

## Future Architecture

The planned evolution of the architecture includes:

1. **Plugin System**:
   - A more robust plugin architecture for different document types
   - Registration mechanism for new converters
   - Dynamic selection of appropriate converter

2. **Unified PowerShell Module**:
   - All converters integrated into a single module
   - Common parameter handling
   - Shared utilities

3. **Enhanced Document Analysis**:
   - More sophisticated parsing of document structure
   - Better handling of various text formats
   - Support for complex nested elements

## Code Design Principles

The codebase follows these design principles:

1. **Clarity over cleverness**: Code is written to be understandable
2. **Proper error handling**: Try/catch blocks with appropriate cleanup
3. **Parameter validation**: Input parameters are validated
4. **Meaningful logging**: Informative messages for troubleshooting
5. **Consistent naming**: PowerShell verb-noun convention for functions

## Integration Points

The scripts integrate with:
- **Word COM Interface**: For creating and formatting documents
- **PowerShell Environment**: For script execution and parameter handling
- **File System**: For reading input and writing output

## Conclusion

The md_txt-to_docx architecture provides a flexible, extensible framework for converting text files to Word documents with intelligent formatting. The modular design allows for future enhancements while maintaining backward compatibility.


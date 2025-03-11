# Document Converter Consolidation Plan

## Current Status

We currently have multiple document converters in our repository:

1. **SemanticDocConverter.ps1** - General-purpose semantic document converter
2. **TrustDocConverter.ps1** - Specialized converter for trust inventory documents
3. **Convert-MarkdownToWord.ps1** - Original converter with limited capabilities

## Phase 1: Code Refinement (Completed)

- ✅ Fixed linter errors in both new converters
- ✅ Updated string interpolation to avoid colon issues
- ✅ Applied proper PowerShell best practices (CmdletBinding, verbose output)
- ✅ Improved table detection and processing
- ✅ Created comprehensive README.md

## Phase 2: Consolidation (Next Steps)

- [ ] Create a unified PowerShell module structure
- [ ] Refactor common code into shared functions
- [ ] Implement plugin architecture for specialized document types
- [ ] Add unit tests for core functionality
- [ ] Create comprehensive documentation for all functions

## Phase 3: Feature Expansion

- [ ] Add support for more table formats (grid tables, CSV data)
- [ ] Implement styling options for different document types
- [ ] Create document templates for common outputs
- [ ] Add image support for Markdown files with image references
- [ ] Implement headers and footers
- [ ] Add automatic table of contents generation

## Phase 4: Distribution & Packaging

- [ ] Package as a PowerShell module for easy installation
- [ ] Create installer script for dependencies
- [ ] Add version checking and update notifications
- [ ] Create examples folder with sample documents
- [ ] Add comprehensive usage guide with examples

## Decision Points

1. Should we maintain backward compatibility with the original converter?
2. Should we create a unified interface or keep specialized converters separate?
3. What additional document types should we support next?
4. Should we use a configuration file for default settings?

## Resources & References

- [PowerShell Best Practices](https://docs.microsoft.com/en-us/powershell/scripting/developer/cmdlet/cmdlet-overview)
- [Word Automation Documentation](https://docs.microsoft.com/en-us/office/vba/api/overview/word)
- [Markdown Syntax Guide](https://www.markdownguide.org/basic-syntax/) 
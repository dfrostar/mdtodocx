# GitHub Wiki Plan for mdtodocx

This document outlines the structure and content for setting up the GitHub Wiki for the mdtodocx repository.

## Wiki Pages Structure

1. **Home**
   - Overview of the tool
   - Quick navigation links to other wiki pages
   - Latest updates/changelog

2. **Installation**
   - Prerequisites and requirements
   - Step-by-step installation instructions
   - Verifying installation

3. **Usage Guide**
   - Basic usage examples
   - Command-line parameters explanation
   - Batch file usage

4. **File Format Support**
   - Markdown (.md) file format details
   - Text (.txt) file format details
   - Table format requirements
   - Supported formatting features

5. **Examples**
   - Markdown to Word examples with screenshots
   - TXT to Word examples with screenshots
   - Example files (with links to the examples directory)

6. **Troubleshooting**
   - Common issues and solutions
   - Error messages explained
   - Word application interaction issues

7. **Advanced Usage**
   - Integrating with other tools
   - Automation scenarios
   - Customizing output formatting

8. **Contributing**
   - How to contribute to the project
   - Development setup
   - Testing guidelines

## Setting Up the Wiki

### Step 1: Enable Wiki in GitHub Repository Settings
1. Go to the repository settings
2. Under "Features," make sure "Wikis" is enabled

### Step 2: Create Wiki Pages
1. Go to the Wiki tab in your repository
2. Click "Create the first page"
3. Use the Home.md template below for the first page
4. Create the remaining pages using the structure above

### Step 3: Link Wiki in README
1. Add a link to the wiki in the README.md
2. Reference specific wiki pages for detailed topics

## Home Page Template

```markdown
# Welcome to the mdtodocx Wiki

This tool allows you to convert Markdown (.md) and plain text (.txt) files to properly formatted Microsoft Word (.docx) documents.

## Quick Navigation

- [Installation](./Installation)
- [Usage Guide](./Usage-Guide)
- [File Format Support](./File-Format-Support)
- [Examples](./Examples)
- [Troubleshooting](./Troubleshooting)
- [Advanced Usage](./Advanced-Usage)
- [Contributing](./Contributing)

## Overview

The mdtodocx tool is a PowerShell-based converter that transforms Markdown and text files into well-formatted Word documents. It preserves formatting, handles tables, lists, and basic text styling.

## Features

- Converts both Markdown (.md) and text (.txt) files
- Preserves tables, headings, and lists
- Maintains text formatting (bold, italic)
- Easy to use through PowerShell or batch file wrapper
- Customizable input and output paths

## Latest Updates

- Added support for TXT files with tables
- Improved table formatting
- Updated batch file wrapper
```

## Implementation Timeline

1. **Week 1**: Create Home, Installation, and Usage Guide pages
2. **Week 2**: Add File Format Support and Examples pages
3. **Week 3**: Add Troubleshooting and Advanced Usage pages
4. **Week 4**: Add Contributing page and refine all content

## Wiki Maintenance Guidelines

- Update the wiki when adding new features
- Include screenshots for visual guidance
- Keep code examples up to date
- Link to specific examples in the repository 
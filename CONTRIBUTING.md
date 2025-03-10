# Contributing to Markdown to Word Converter

Thank you for considering contributing to the Markdown to Word Converter! This document provides guidelines and instructions for contributing to this project.

## Code of Conduct

By participating in this project, you are expected to uphold our Code of Conduct:

- Be respectful and inclusive
- Be patient and welcoming
- Be thoughtful in your criticism
- Focus on what is best for the community

## How Can I Contribute?

### Reporting Bugs

This section guides you through submitting a bug report. Following these guidelines helps maintainers understand your report, reproduce the behavior, and find related reports.

**Before Submitting A Bug Report:**

- Check the documentation for tips on using the script correctly
- Check if the issue has already been reported in the Issues section
- Try different parameters or configurations to see if the issue is consistent

**How Do I Submit A Good Bug Report?**

Bugs are tracked as GitHub issues. Create an issue and provide the following information:

- **Use a clear and descriptive title** for the issue
- **Describe the exact steps to reproduce the problem** with as much detail as possible
- **Provide specific examples** such as the exact command line used
- **Describe the behavior you observed** and point out why it's problematic
- **Include error messages and stack traces** that appeared in the PowerShell console
- **Specify your PowerShell version, Word version, and OS**
- If possible, include screenshots or animated GIFs

### Suggesting Enhancements

Enhancement suggestions are also tracked as GitHub issues. When creating an enhancement suggestion, please include:

- **Use a clear and descriptive title** for the issue
- **Provide a step-by-step description of the suggested enhancement**
- **Provide specific examples** to demonstrate the proposed enhancement
- **Describe the current behavior** and **explain how this enhancement would be useful**
- **List other applications where this enhancement exists**, if applicable

### Pull Requests

- Fill in the required template
- Do not include issue numbers in the PR title
- Include screenshots and animated GIFs if possible
- Follow the PowerShell style guide
- Include adequate tests
- Document new features in the README or documentation
- Update the version number following [Semantic Versioning](https://semver.org/)

## Style Guides

### Git Commit Messages

- Use the present tense ("Add feature" not "Added feature")
- Use the imperative mood ("Move cursor to..." not "Moves cursor to...")
- Limit the first line to 72 characters or less
- Reference issues and pull requests liberally after the first line

### PowerShell Style Guide

- Use camelCase for variables and PascalCase for function names
- Include clear, comprehensive comments
- Add help documentation for functions using Comment-Based Help
- Use standard PowerShell cmdlet verbs (Get-, Set-, New-, Remove-, etc.)
- Prefer explicit parameter names over positional parameters

## Development Process

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Make your changes
4. Test your changes thoroughly
5. Commit your changes (`git commit -m 'Add some amazing feature'`)
6. Push to the branch (`git push origin feature/amazing-feature`)
7. Open a Pull Request

## Testing

- Test any changes with various Markdown files
- Verify that the conversion works correctly for:
  - Tables
  - Lists
  - Headings
  - Text formatting
- Ensure no regression in existing functionality

## Documentation

- Update the README.md file with details of changes to the interface
- Update the USAGE.md file with any new parameters or usage examples
- Comment your code thoroughly

Thank you for your contribution! 
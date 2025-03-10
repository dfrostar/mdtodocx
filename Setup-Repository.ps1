<#
.SYNOPSIS
    Sets up the directory structure and files for the Markdown to Word Converter GitHub repository.

.DESCRIPTION
    This script creates the necessary directories and files for a GitHub repository
    containing the Markdown to Word Converter PowerShell script. It sets up documentation,
    examples, and the core script file.

.PARAMETER RepositoryPath
    The path where the repository structure will be created. Default is the current directory.

.PARAMETER UserName
    Your GitHub username to be used in documentation. Default is "YourUsername".

.PARAMETER RepoName
    The name of the repository. Default is "markdown-to-word".

.EXAMPLE
    .\Setup-Repository.ps1 -RepositoryPath "C:\Projects\markdown-to-word" -UserName "johndoe" -RepoName "md-to-word"

.NOTES
    Requires PowerShell 5.1 or higher.
#>

param (
    [string]$RepositoryPath = (Get-Location),
    [string]$UserName = "YourUsername",
    [string]$RepoName = "markdown-to-word"
)

# Create the directory structure
function Create-DirectoryStructure {
    param (
        [string]$BasePath
    )

    Write-Host "Creating directory structure..." -ForegroundColor Cyan

    # Main directories
    $directories = @(
        "$BasePath\examples",
        "$BasePath\docs",
        "$BasePath\tests"
    )

    foreach ($dir in $directories) {
        if (-not (Test-Path $dir)) {
            New-Item -Path $dir -ItemType Directory -Force | Out-Null
            Write-Host "  Created: $dir" -ForegroundColor Green
        }
        else {
            Write-Host "  Directory already exists: $dir" -ForegroundColor Yellow
        }
    }
}

# Copy files to their locations
function Copy-Files {
    param (
        [string]$BasePath,
        [string]$UserName,
        [string]$RepoName
    )

    Write-Host "Setting up repository files..." -ForegroundColor Cyan

    # List of files to copy
    $files = @{
        "Convert-MarkdownToWord.ps1" = "$BasePath\Convert-MarkdownToWord.ps1"
        "README.md" = "$BasePath\README.md"
        "LICENSE" = "$BasePath\LICENSE"
        "CONTRIBUTING.md" = "$BasePath\CONTRIBUTING.md"
        "USAGE.md" = "$BasePath\docs\USAGE.md"
        "sample.md" = "$BasePath\examples\sample.md"
    }

    foreach ($file in $files.Keys) {
        $sourcePath = Join-Path (Get-Location) $file
        $destPath = $files[$file]
        
        if (Test-Path $sourcePath) {
            Copy-Item -Path $sourcePath -Destination $destPath -Force
            Write-Host "  Copied: $file to $destPath" -ForegroundColor Green
            
            # Update GitHub username references in the README and USAGE files
            if ($file -eq "README.md" -or $file -eq "USAGE.md") {
                $content = Get-Content -Path $destPath -Raw
                $content = $content -replace "yourusername", $UserName -replace "markdown-to-word", $RepoName
                Set-Content -Path $destPath -Value $content
                Write-Host "  Updated references in $file" -ForegroundColor Green
            }
        }
        else {
            Write-Host "  Warning: Source file not found: $sourcePath" -ForegroundColor Yellow
        }
    }
}

# Create .gitignore file
function Create-GitIgnore {
    param (
        [string]$BasePath
    )

    $gitignorePath = Join-Path $BasePath ".gitignore"
    $gitignoreContent = @"
# PowerShell artifacts
*.ps1xml
*.psm1
*.psd1
*.log
*.ps1.bak

# PowerShell module dependencies
/modules/

# Windows artifacts
Thumbs.db
desktop.ini

# macOS artifacts
.DS_Store

# Temporary files
*.tmp
*~

# Word temporary files
~$*.docx
~$*.doc

# IDE files
.vscode/
.idea/
*.suo
*.user
*.userosscache
*.sln.docstates
"@

    Set-Content -Path $gitignorePath -Value $gitignoreContent
    Write-Host "  Created: .gitignore" -ForegroundColor Green
}

# Create a tests directory with a basic test script
function Create-TestScript {
    param (
        [string]$BasePath
    )

    $testScriptPath = Join-Path $BasePath "tests\Test-Conversion.ps1"
    $testScriptContent = @"
<#
.SYNOPSIS
    Tests the Convert-MarkdownToWord.ps1 script with various input files.

.DESCRIPTION
    This script tests the Markdown to Word converter with different sample files
    to ensure functionality is working as expected.

.EXAMPLE
    .\Test-Conversion.ps1
#>

# Get the root directory of the repository
`$scriptPath = Split-Path -Parent `$MyInvocation.MyCommand.Path
`$repoRoot = Split-Path -Parent `$scriptPath

# Source the main script
`$converterScript = Join-Path `$repoRoot "Convert-MarkdownToWord.ps1"

if (-not (Test-Path `$converterScript)) {
    Write-Error "Converter script not found at: `$converterScript"
    exit 1
}

# Define test cases
`$testCases = @(
    @{
        Name = "Simple Table Test"
        InputFile = Join-Path `$repoRoot "examples\sample.md"
        OutputFile = Join-Path `$scriptPath "output\sample-output.docx"
    }
    # Add more test cases as needed
)

# Create output directory if it doesn't exist
`$outputDir = Join-Path `$scriptPath "output"
if (-not (Test-Path `$outputDir)) {
    New-Item -Path `$outputDir -ItemType Directory -Force | Out-Null
}

# Run tests
foreach (`$test in `$testCases) {
    Write-Host "Running test: `$(`$test.Name)" -ForegroundColor Cyan
    
    # Ensure input file exists
    if (-not (Test-Path `$test.InputFile)) {
        Write-Error "Input file not found: `$(`$test.InputFile)"
        continue
    }
    
    # Run the converter
    try {
        & `$converterScript -InputFile `$test.InputFile -OutputFile `$test.OutputFile
        
        if (Test-Path `$test.OutputFile) {
            Write-Host "  Test successful: Output file created at `$(`$test.OutputFile)" -ForegroundColor Green
        }
        else {
            Write-Host "  Test failed: Output file not created" -ForegroundColor Red
        }
    }
    catch {
        Write-Host "  Test failed with error: `$_" -ForegroundColor Red
    }
}

Write-Host "All tests completed" -ForegroundColor Cyan
"@

    $testScriptDir = Join-Path $BasePath "tests"
    if (-not (Test-Path $testScriptDir)) {
        New-Item -Path $testScriptDir -ItemType Directory -Force | Out-Null
    }

    Set-Content -Path $testScriptPath -Value $testScriptContent
    Write-Host "  Created: Test script at $testScriptPath" -ForegroundColor Green
}

# Main function to set up the repository
function Setup-Repository {
    param (
        [string]$RepositoryPath,
        [string]$UserName,
        [string]$RepoName
    )

    Write-Host "Setting up Markdown to Word Converter repository at: $RepositoryPath" -ForegroundColor Cyan
    
    # Create base directory if it doesn't exist
    if (-not (Test-Path $RepositoryPath)) {
        New-Item -Path $RepositoryPath -ItemType Directory -Force | Out-Null
        Write-Host "Created repository directory: $RepositoryPath" -ForegroundColor Green
    }
    
    # Create directory structure
    Create-DirectoryStructure -BasePath $RepositoryPath
    
    # Copy files
    Copy-Files -BasePath $RepositoryPath -UserName $UserName -RepoName $RepoName
    
    # Create .gitignore
    Create-GitIgnore -BasePath $RepositoryPath
    
    # Create test script
    Create-TestScript -BasePath $RepositoryPath
    
    Write-Host "Repository setup complete!" -ForegroundColor Green
    Write-Host "Next steps:" -ForegroundColor Cyan
    Write-Host "1. Navigate to $RepositoryPath" -ForegroundColor White
    Write-Host "2. Initialize Git repository with: git init" -ForegroundColor White
    Write-Host "3. Add files with: git add ." -ForegroundColor White
    Write-Host "4. Commit with: git commit -m 'Initial commit'" -ForegroundColor White
    Write-Host "5. Create a new repository on GitHub" -ForegroundColor White
    Write-Host "6. Link and push with: git remote add origin https://github.com/$UserName/$RepoName.git && git push -u origin master" -ForegroundColor White
}

# Run the setup
Setup-Repository -RepositoryPath $RepositoryPath -UserName $UserName -RepoName $RepoName 
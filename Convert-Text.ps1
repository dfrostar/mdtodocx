[CmdletBinding()]
param(
    [Parameter(Mandatory=$true, Position=0)]
    [string]$InputFile,
    
    [Parameter(Mandatory=$false, Position=1)]
    [string]$OutputFile,
    
    [Parameter(Mandatory=$false)]
    [switch]$ShowWord
)

# Set default output file if not provided
if ([string]::IsNullOrEmpty($OutputFile)) {
    $OutputFile = [System.IO.Path]::ChangeExtension($InputFile, "docx")
}

Write-Host "Converting: $InputFile to $OutputFile"
if ($ShowWord) {
    Write-Host "Word will be visible during conversion"
}

# Get the directory of the current script
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Run the TrustInventory script (which works well for general text too)
$params = @{
    InputFile = $InputFile
    OutputFile = $OutputFile
}

if ($ShowWord) {
    $params.Add("ShowWord", $true)
}

try {
    & "$scriptDir\Test-TrustInventory.ps1" @params
    
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Conversion failed with error code $LASTEXITCODE"
        exit $LASTEXITCODE
    }
    
    Write-Host "`nConversion completed successfully!"
    Write-Host "Output saved to: $OutputFile`n"
    exit 0
}
catch {
    Write-Error "An error occurred during conversion: $_"
    exit 1
} 
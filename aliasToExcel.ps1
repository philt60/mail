# Enable strict mode and verbose output
Set-StrictMode -Version Latest
$VerbosePreference = "Continue"

# Import the ImportExcel module if it's not already loaded
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Import-Module -Name ImportExcel
}

# Define the path to the input file
$inputFilePath = Join-Path -Path $PSScriptRoot -ChildPath "aliases" -Resolve

# Read the data from the file
$lines = Get-Content -Path $inputFilePath

# Initialize an array to hold the parsed data
$parsedData = @()

# Parse each line
foreach ($line in $lines) {
    $line = $line.Trim()
    if (-not [string]::IsNullOrWhiteSpace($line) -and -not $line.StartsWith("#")) {
        if ($line -match "^(?<Name>.*?):\s*(?<Emails>.*)$") {
            $name = $matches['Name'].Trim()
            $emails = $matches['Emails'].Trim().TrimEnd(',')
            $emailNames = ($emails -split ",") | ForEach-Object {
                if ($_ -match "@mit.edu$") {
                    ($_ -split "@mit.edu")
                } else {
                    $_
                }
            } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            $parsedData += [PSCustomObject]@{
                Name = $name
                Emails = $emails
                EmailNames = ($emailNames -join ", ")
            }
        }
    }
}

# Export the data to an Excel file and show it
# $parsedData | Export-Csv -Path "aliases.csv" -NoTypeInformation
# $parsedData | Export-Excel -Path "C:\temp\aliases.xlsx" -WorksheetName "Data" -AutoSize -TableStyle "Medium16" -Show -Verbose
$parsedData | Out-GridView




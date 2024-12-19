# Phil's script to overlay csv files.

Set-StrictMode -Version Latest
$FormatEnumerationLimit = -1

Function OverlayCSVs
{
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param (
        [Parameter(Mandatory=$false)]
        [ValidateScript({
            if (-Not (Test-Path -Path $_ -PathType Leaf) ) {
                throw "Cannot find input file."
            }
            if ($_ -notmatch "\.csv$") {
                throw "The file must be .csv"
            }
            return $true
        })]
        [string]$inputFile1 = "mail7_alias_counts2023.csv",
        [Parameter(Mandatory=$false)]
        [ValidateScript({
            if (-Not (Test-Path -Path $_ -PathType Leaf) ) {
                throw "Cannot find input file."
            }
            if ($_ -notmatch "\.csv$") {
                throw "The file must be .csv"
            }
            return $true
        })]
        [string]$inputFile2 = "mail8_alias_counts2023.csv",
        [Parameter(Mandatory=$false)]
        [ValidateScript({
            if (-Not (Test-Path -Path $_ -PathType Leaf -IsValid) ) {
                throw "Output path is not valid."
            }
            if ($_ -notmatch "\.csv$") {
                throw "The file must be .csv"
            }
            return $true
        })]
        [string]$outputFile = "combined_all2024.csv"
        # [string]$outputFile = "test_out.csv"
    )

    Write-Verbose "inputFiles: $inputFile1 $inputFile2  outputFile: $outputFile"
    if ($inputFile1 -eq $inputFile2 -or $inputFile1 -eq $outputFile) {
        Write-Warning "All files must be separate: inputFile1, $inputFile2, $outputFile"
        return exit 10
    }
    try {
        $myCsv1 = Import-Csv -LiteralPath $inputFile1 -ErrorAction Stop
        $myCsv2 = Import-Csv -LiteralPath $inputFile2 -ErrorAction Stop
    } catch {
        Write-Warning "$_"
        exit 20
    }
    # $colNames = ($myCsv | Get-Member -MemberType NoteProperty).Name   # alphabetically sorted list :-(
    $colNames1 = $myCsv1[0].psobject.Properties.Name
    $colNames2 = $myCsv2[0].psobject.Properties.Name
    if (Compare-Object -ReferenceObject $colNames1 -DifferenceObject $colNames2) {
        Write-Warning("Column names don't match")
        exit 90
    }
    if ($myCsv1.Count -ne $myCsv2.Count) {
        Write-Warning("Row counts don't match {0} : {1}" -f $myCsv1.Count, $myCsv2.Count)
        exit 92
    }
    Write-Verbose("CSV counts columns:{0} rows:{1}" -f $colNames1.Count, $myCsv1.Count)

    for ($i = 0; $i -lt $myCsv1.count; $i++) {
        foreach ($cName in $colNames1) {
            if ($cName -match "date") {
                if ($myCsv1[$i].$cName -ne $myCsv2[$i].$cName) {
                    Write-Warning("Dates don't mactch {0} != {1} at line {3}" -f $myCsv1[$i].$cName, $myCsv2[$i].$cName, $i)
                    exit 93
                }
                continue
            }
            if ([string]::IsNullOrEmpty($myCsv1[$i].$cName) -and [string]::IsNullOrEmpty($myCsv2[$i].$cName)) {
                $myCsv1[$i].$cName = $null
            } else {
                $myCsv1[$i].$cName = [int]$myCsv1[$i].$cName + [int]$myCsv2[$i].$cName
            }
        }
    }

    Remove-Item -LiteralPath $outputFile -ErrorAction SilentlyContinue
    try {
        if ($PSVersionTable.PSVersion.Major -eq 5) {
            $myCsv1 | Export-Csv -LiteralPath $outputFile -NoTypeInformation -UseCulture -Encoding ASCII -ErrorAction Stop
        } else {
            $myCsv1 | Export-Csv -LiteralPath $outputFile -NoTypeInformation -UseCulture -Encoding ASCII -UseQuotes AsNeeded -ErrorAction Stop
        }
    } catch {
        Write-Warning "$_"
    }
}

# main
Set-Location $PSScriptRoot
OverlayCSVs -inputFile1 "mail7_alias_counts2023.csv" -inputFile2 "mail8_alias_counts2023.csv" -outputFile "combined_all2024.csv" -Verbose
# eof
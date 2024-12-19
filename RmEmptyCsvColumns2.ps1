# Phil's script to drop empty columns from a csv file.

Set-StrictMode -Version Latest
$FormatEnumerationLimit = -1

Function RemoveEmptyColumns
{
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param (
        [Parameter(Mandatory=$false)]
        [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
        [ValidatePattern('\.csv')]
        [string]$inputFile = "testIn.csv",
        [Parameter(Mandatory=$false)]
        [ValidateScript({ Test-Path -Path $_ -PathType Leaf -IsValid })]
        [ValidatePattern('\.csv')]
        [string]$outputFile = "testOut.csv",
        [switch]$listing                            # list columns kept/dropped to stdout
    )

    Write-Verbose "inputFile: $inputFile found  outputFile: $outputFile"
    if ($inputFile -eq $outputFile) {
        Write-Warning "Cannot overwrite input file ($inputFile -eq $outputFile). Bye"
        exit 10
    }
    try {
        if (!($myCsv = Import-Csv -LiteralPath $inputFile -ErrorAction Stop)) {
            Write-Warning "No data returned in $inputFile. Bye"
            exit 11
        }
    } catch {
        Write-Warning "$_"
        exit 20
    }
        
    # $colNames = ($myCsv | Get-Member -MemberType NoteProperty).Name   # alphabetically sorted list :-(
    $colNames = @($myCsv[0].psobject.Properties.Name)
    if ($colNames.Count -le 1) {
        Write-Warning ("Column count looks weird {0}. Bye" -f $colNames.Count)
        exit 21
    }
    Write-Verbose ("CSV counts columns: {0}  rows:{1}" -f $colNames.Count, $myCsv.Count)
    $dataCols = @(); $droppedCols = @();
    foreach ($cName in $colNames) {
        if (@($myCsv.$cName | ? {$_}).Count -ne 0) {   # go to each column and count where values (exist)
            $dataCols += $cName 
        } else {
            $droppedCols += $cName
        }
    }
    if ($PSBoundParameters['listing']) {
        "Columns kept ({0}):" -f $dataCols.Count
        $dataCols.foreach({$_})
        "`nColumns dropped ({0}):" -f $droppedCols.Count
        $droppedCols.foreach({$_})
    }
    Remove-Item -LiteralPath $outputFile -Force -ErrorAction SilentlyContinue
    try {
        if ($PSVersionTable.PSVersion.Major -eq 7) {
            $myCsv | Select-Object -Property $dataCols | Export-Csv -LiteralPath $outputFile -NoTypeInformation -UseCulture -Encoding ASCII -UseQuotes AsNeeded -ErrorAction Stop
        } else {
            $myCsv | Select-Object -Property $dataCols | Export-Csv -LiteralPath $outputFile -NoTypeInformation -UseCulture -Encoding ASCII -ErrorAction Stop
        }
    } catch {
        Write-Warning "$_"
        exit 30
    }
}

# main
Add-Type -AssemblyName System.Windows.Forms
$inFileDlg = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('Desktop') 
    Filter = 'SpreadSheet (*.csv)|*.csv'
}
if ($inFileDlg.ShowDialog() -ne "OK") {
    Write-Host "No input file selected"
    exit 40
}

$inFileDlg | fl *
Set-Location (Split-Path -LiteralPath $inFileDlg.FileName)
# RemoveEmptyColumns -inputFile test1a.csv -outputFile testOut.csv -listing -Verbose # -listing
RemoveEmptyColumns -inputFile $inFileDlg.FileName -outputFile testOut.csv -listing -Verbose
$inFileDlg.Dispose()
# eof
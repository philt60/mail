# Import the required module
Import-Module ImportExcel
Set-Location $PSScriptRoot

# Import the data from both Excel files
$members = Import-Excel -Path ".\Members.xlsx" -WorksheetName 'Members'
$combined = Import-Excel -Path ".\combined_all2024.xlsx"

# Get the totals from the "Total" row of combined_all2024 2.xlsx
$totalsRow = $combined | Where-Object Date -eq "Total"

# Loop through each row in Members.xlsx
foreach ($row in $members) {
    # Find the matching total in the totals row based on the "Name" column
    $total = $totalsRow.PSObject.Properties[$row.Name]
    # $total = $totalsRow.($row.Name)

    # If a matching total is found, add it to the member row
    if ($total) {
        $row | Add-Member -NotePropertyName 'Total' -NotePropertyValue $total.Value
    }
}

# Check if the worksheet exists and remove it if it does
$worksheetName = "totals"
if (Test-Path ".\Members.xlsx") {
    $excel = Open-ExcelPackage -Path ".\Members.xlsx"
    $worksheet = $excel.Workbook.Worksheets[$worksheetName]
    if ($worksheet) {
        $excel.Workbook.Worksheets.Delete($worksheetName)
    }
    Close-ExcelPackage $excel
}
# Reorder columns and export the updated data back to Members.xlsx
$members | Select-Object Name, Email, Total, Class, Members | Export-Excel -Path ".\Members.xlsx" -WorksheetName $worksheetName -AutoSize -TableStyle Medium16 -FreezeTopRowFirstColumn -MoveToStart -Show
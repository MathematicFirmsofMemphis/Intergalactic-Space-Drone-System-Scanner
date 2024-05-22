# Scanner.ps1

# Define the Excel application and workbook
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$Workbook = $Excel.Workbooks.Add()
$Worksheet = $Workbook.Worksheets.Item(1)

# Define header names and sample data for EVE Online account
$Headers = @("Character Name", "Skill Points", "Credit Balance", "Corporation", "Ship Type")
$Data = @(
    @("John Doe", 2000000, 5000000, "Corp A", "Space Drone-Frigate"),
    @("Jane Smith", 3500000, 12000000, "Corp B", "Space Drone-Cruiser"),
    @("Max Mustermann", 5000000, 3000000, "Corp C", "Space Drone-Destroyer")
)

# Write headers to the first row
for ($i = 0; $i -lt $Headers.Length; $i++) {
    $Worksheet.Cells.Item(1, $i + 1) = $Headers[$i]
}

# Write data to subsequent rows
for ($row = 0; $row -lt $Data.Length; $row++) {
    for ($col = 0; $col -lt $Data[$row].Length; $col++) {
        $Worksheet.Cells.Item($row + 2, $col + 1) = $Data[$row][$col]
    }
}

# Adjust column width
$Worksheet.Columns.AutoFit()

# Save the Excel workbook
$OutputFile = "$PSScriptRoot\ScannerData.xlsx"
$Workbook.SaveAs($OutputFile)
$Excel.Quit()

# Release COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Worksheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null

# Output file path
Write-Output "Excel file created at: $OutputFile"

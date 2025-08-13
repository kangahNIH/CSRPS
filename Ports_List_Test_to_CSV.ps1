# Define input Excel path and output CSV path
$excelPath = "C:\STPS\ports-source-destination.csv"
$outputCsv = "C:\STPS\PortTestResults.csv"

# Load Excel COM Object
$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Open($excelPath)
$sheet = $workbook.Sheets.Item(1)

# Initialize result array
$results = @()

# Read rows from 2 to last used row
$rowCount = $sheet.UsedRange.Rows.Count

# Destination columns E to J (columns 5 to 10)
$destinationCols = 5..10
$sourceIP = "10.181.27.52"

for ($row = 2; $row -le $rowCount; $row++) {
    $port = $sheet.Cells.Item($row, 1).Value2

    foreach ($col in $destinationCols) {
        $destination = $sheet.Cells.Item(1, $col).Value2
        if ($destination -ne $null -and $port -ne $null) {
            Write-Host "Testing $destination on port $port from $sourceIP..."
            $test = Test-NetConnection -ComputerName $destination -Port $port -InformationLevel Quiet

            $results += [PSCustomObject]@{
                Source       = $sourceIP
                Destination  = $destination
                Port         = $port
                Status       = if ($test) { "Open" } else { "Closed" }
            }
        }
    }
}

# Cleanup
$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

# Export to CSV
$results | Export-Csv -Path $outputCsv -NoTypeInformation -Encoding UTF8
Write-Host "Results exported to $outputCsv"

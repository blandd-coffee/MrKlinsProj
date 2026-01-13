# Extract data from Excel and create JSON files
$excelPath = 'C:\Users\CMP_AnSpencer\Desktop\Projects\MrKlinsProj\public\RegionalContracts.xlsx'
$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Open($excelPath)

$schools = @(
    @{ SheetName = "Erie School District 26-27"; FileName = "Erie.json" }
    @{ SheetName = "North East 26-27"; FileName = "NorthEast.json" }
)

foreach ($school in $schools) {
    $sheet = $workbook.Sheets.Item($school.SheetName)
    $usedRange = $sheet.UsedRange
    $rows = $usedRange.Rows.Count
    $cols = $usedRange.Columns.Count
    
    # Get headers (first row)
    $headers = @()
    for ($c = 1; $c -le $cols; $c++) {
        $header = $sheet.Cells.Item(1, $c).Value
        if ($header) {
            $headers += $header
        }
    }
    
    # Extract data rows
    $data = @()
    for ($r = 2; $r -le $rows; $r++) {
        $rowObj = @{}
        $hasData = $false
        
        for ($c = 1; $c -le $cols; $c++) {
            $header = $headers[$c - 1]
            $value = $sheet.Cells.Item($r, $c).Value
            
            if ($header) {
                # Only include numeric values or "step" column
                if ($header -eq "step" -or [double]::TryParse($value, [ref]$null)) {
                    $rowObj[$header] = $value
                    if ($value) { $hasData = $true }
                }
            }
        }
        
        # Only add row if it has data
        if ($hasData) {
            $data += $rowObj
        }
    }
    
    # Convert to JSON and save
    $jsonPath = "C:\Users\CMP_AnSpencer\Desktop\Projects\MrKlinsProj\public\$($school.FileName)"
    $data | ConvertTo-Json -Depth 10 | Out-File -FilePath $jsonPath -Encoding UTF8
    Write-Host "Created $($school.FileName) with $($data.Count) rows"
}

$workbook.Close()
$excel.Quit()

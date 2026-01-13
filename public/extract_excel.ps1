# Extract data from Excel and create JSON files
$excelPath = 'C:\Users\CMP_AnSpencer\Desktop\Projects\MrKlinsProj\public\RegionalContracts.xlsx'

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false  # Run in background to avoid UI freezing
    $workbook = $excel.Workbooks.Open($excelPath)

    $schools = @(
        @{ SheetName = "Crawford Central School 26-27"; FileName = "Crawford.json" }
        @{ SheetName = "Corry 23-24"; FileName = "Corry.json" }
        @{ SheetName = "Mercer 26-27"; FileName = "Mercer.json" }
        @{ SheetName = "Erie School District 26-27"; FileName = "Erie.json" }
        @{ SheetName = "North East 26-27"; FileName = "NorthEast.json" }
        @{ SheetName = "ECTS 26-27"; FileName = "ECTS.json" }
    )

    foreach ($school in $schools) {
        try {
            $sheet = $workbook.Sheets.Item($school.SheetName)
            $usedRange = $sheet.UsedRange
            $rows = $usedRange.Rows.Count
            $cols = $usedRange.Columns.Count

            # Get headers (first row) - preserve column alignment
            $headers = @()
            for ($c = 1; $c -le $cols; $c++) {
                $headers += "$($sheet.Cells.Item(1, $c).Value2)"
            }

            # Extract data rows
            $data = @()
            for ($r = 2; $r -le $rows; $r++) {
                $rowObj = @{}
                $hasData = $false

                for ($c = 1; $c -le $cols; $c++) {
                    $header = [string]$headers[$c - 1]
                    if (-not $header) { continue }

                    $value = $sheet.Cells.Item($r, $c).Value

                    if (
                        $header.ToLower() -eq "step" -or
                        [double]::TryParse("$value", [ref]$null)
                    ) {
                        $rowObj[$header] = $value
                        if ($null -ne $value -and $value -ne "") {
                            $hasData = $true
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
        } catch {
            Write-Host "Error processing $($school.SheetName): $($_.Exception.Message)"
        }
    }

    $workbook.Close($false)
} catch {
    Write-Host "Error opening Excel file: $($_.Exception.Message)"
} finally {
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

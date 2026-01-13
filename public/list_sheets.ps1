$excelPath = 'C:\Users\CMP_AnSpencer\Desktop\Projects\MrKlinsProj\public\RegionalContracts.xlsx'
$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Open($excelPath)
$workbook.Sheets | ForEach-Object { Write-Host $_.Name }
$workbook.Close()
$excel.Quit()

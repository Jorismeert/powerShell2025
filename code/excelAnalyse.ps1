
# Import Excel file and display data
$filePath = "/Users/jorismeert/Desktop/Powershell/project/data/orders_mrt_sep.xlsx"
$data = Import-Excel -Path $filePath
# Basic information about the data
Write-Host "=== BASIC DATA ANALYSIS ===" -ForegroundColor Green
Write-Host "Total rows: $($data.Count)" -ForegroundColor Yellow
Write-Host "Data type: $($data.GetType().Name)" -ForegroundColor Yellow

# Column names and data types
Write-Host "`n=== COLUMN INFORMATION ===" -ForegroundColor Green
$firstRow = $data | Select-Object -First 1
$columnCount = ($firstRow.PSObject.Properties | Measure-Object).Count
Write-Host "Total columns: $columnCount" -ForegroundColor Yellow

Write-Host "`nColumn Names and Sample Values:" -ForegroundColor Cyan
$firstRow.PSObject.Properties | ForEach-Object {
    $columnName = $_.Name
    $sampleValue = $_.Value
    $dataType = if ($sampleValue -ne $null) { $sampleValue.GetType().Name } else { "NULL" }
    Write-Host "  $columnName : $sampleValue ($dataType)" -ForegroundColor White
}
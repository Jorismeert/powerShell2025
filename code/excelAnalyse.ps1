# Complete Comprehensive Data Analysis Script

# Define all required functions first
function Get-DataOverview {
    param($Data)
    
    Write-Host "`nFirst 3 rows:" -ForegroundColor Cyan
    $Data | Select-Object -First 3 | Format-Table -AutoSize
    
    Write-Host "`nLast 3 rows:" -ForegroundColor Cyan
    $Data | Select-Object -Last 3 | Format-Table -AutoSize
    
    Write-Host "`nColumn Summary:" -ForegroundColor Cyan
    $columns = $Data[0].PSObject.Properties.Name
    $columns | ForEach-Object { 
        $col = $_
        $sample = $Data[0].$col
        $type = if ($sample -ne $null) { $sample.GetType().Name } else { "NULL" }
        Write-Host "  $col : $type" 
    }
}

function Analyze-DataColumns {
    param($Data)
    
    Write-Host "`n=== DETAILED COLUMN ANALYSIS ===" -ForegroundColor Green
    
    $columns = $Data[0].PSObject.Properties.Name
    
    foreach ($column in $columns) {
        Write-Host "`n--- Analyzing: $column ---" -ForegroundColor Cyan
        
        $values = $Data.$column
        $nonNullValues = $values | Where-Object { $_ -ne $null -and $_ -ne "" }
        $nonNullCount = $nonNullValues.Count
        $nullCount = $values.Count - $nonNullCount
        
        if ($nonNullCount -gt 0) {
            # Sample values
            $sampleValues = $nonNullValues | Select-Object -First 5
            Write-Host "Sample values: $($sampleValues -join ', ')"
            
            # Data type analysis
            $dataTypes = $nonNullValues | ForEach-Object { $_.GetType().Name } | Group-Object
            Write-Host "Data types: $(($dataTypes | ForEach-Object { "$($_.Name):$($_.Count)" }) -join ', ')"
            
            # Numeric analysis - CORRECTED
            $numericValues = $nonNullValues | Where-Object { 
                try { [double]$_; $true } catch { $false }
            }
            
            if ($numericValues.Count -gt 0) {
                $convertedValues = $numericValues | ForEach-Object { [double]$_ }
                $stats = $convertedValues | Measure-Object -Minimum -Maximum -Average -Sum -StandardDeviation
                Write-Host "Min value: $($stats.Minimum)"
                Write-Host "Max value: $($stats.Maximum)" 
                Write-Host "Average: $([math]::Round($stats.Average, 2))"
                Write-Host "Sum: $([math]::Round($stats.Sum, 2))"
                Write-Host "Standard Deviation: $([math]::Round($stats.StandardDeviation, 2))"
            }
            
            # String analysis
            $stringValues = $nonNullValues | Where-Object { $_ -is [string] }
            if ($stringValues.Count -gt 0) {
                $uniqueStrings = $stringValues | Group-Object
                Write-Host "Unique string values: $($uniqueStrings.Count)"
                if ($uniqueStrings.Count -le 10) {
                    Write-Host "All unique values: $(($uniqueStrings.Name | Sort-Object) -join ', ')"
                }
            }
        }
    }
}

function Test-DataQuality {
    param($Data)
    
    Write-Host "`n=== DATA QUALITY CHECK ===" -ForegroundColor Green
    
    $qualityReport = @()
    $columns = $Data[0].PSObject.Properties.Name
    
    foreach ($column in $columns) {
        $values = $Data.$column
        $totalRows = $values.Count
        
        $nullCount = ($values | Where-Object { $_ -eq $null }).Count
        $emptyCount = ($values | Where-Object { $_ -eq "" }).Count
        $totalNullOrEmpty = $nullCount + $emptyCount
        
        $dataTypes = $values | Where-Object { $_ -ne $null } | ForEach-Object { $_.GetType().Name } | Group-Object
        $hasMultipleTypes = $dataTypes.Count -gt 1
        
        $qualityReport += [PSCustomObject]@{
            ColumnName = $column
            TotalRows = $totalRows
            NullValues = $nullCount
            EmptyValues = $emptyCount
            NullOrEmptyPercentage = [math]::Round(($totalNullOrEmpty / $totalRows) * 100, 2)
            DataTypes = ($dataTypes.Name -join ', ')
            HasMultipleTypes = $hasMultipleTypes
        }
    }
    
    $qualityReport | Format-Table -AutoSize
    
    Write-Host "`n=== DATA QUALITY SUMMARY ===" -ForegroundColor Yellow
    $problemColumns = $qualityReport | Where-Object { $_.NullOrEmptyPercentage -gt 50 -or $_.HasMultipleTypes }
    if ($problemColumns) {
        Write-Host "Potential issues found in:" -ForegroundColor Red
        $problemColumns | ForEach-Object { 
            Write-Host "  - $($_.ColumnName) ($($_.NullOrEmptyPercentage)% null/empty)" -ForegroundColor Red 
        }
    } else {
        Write-Host "No major data quality issues detected." -ForegroundColor Green
    }
}

function Get-StatisticalSummary {
    param($Data)
    
    Write-Host "`n=== STATISTICAL SUMMARY ===" -ForegroundColor Green
    
    $numericColumns = $Data[0].PSObject.Properties | 
        Where-Object { $_.Value -is [int] -or $_.Value -is [double] -or $_.Value -is [decimal] } |
        Select-Object -ExpandProperty Name
    
    if ($numericColumns) {
        foreach ($column in $numericColumns) {
            Write-Host "`n--- $column ---" -ForegroundColor Yellow
            $numericValues = $Data.$column | Where-Object { $_ -is [int] -or $_ -is [double] -or $_ -is [decimal] }
            
            if ($numericValues.Count -gt 0) {
                $stats = $numericValues | Measure-Object -Minimum -Maximum -Average -Sum -StandardDeviation
                [PSCustomObject]@{
                    Count = $stats.Count
                    Min = $stats.Minimum
                    Max = $stats.Maximum
                    Average = [math]::Round($stats.Average, 2)
                    Sum = $stats.Sum
                    StdDev = [math]::Round($stats.StandardDeviation, 2)
                } | Format-Table -AutoSize
            } else {
                Write-Host "No numeric values found in this column" -ForegroundColor Gray
            }
        }
    } else {
        Write-Host "No numeric columns found for statistical analysis." -ForegroundColor Yellow
    }
}

# Main Analysis Execution
Write-Host "COMPREHENSIVE DATA ANALYSIS" -ForegroundColor Magenta
Write-Host "===========================" -ForegroundColor Magenta

# Check if data exists
if (-not $data -or $data.Count -eq 0) {
    Write-Host "ERROR: No data found. Please import your Excel file first." -ForegroundColor Red
    Write-Host "Use: `$data = Import-Excel -Path `"/Users/jorismeert/Desktop/Powershell/project/data/Beschikbaarheid_Geel_18062025.xlsx`"" -ForegroundColor Yellow
    exit
}

# 1. Basic info
Write-Host "`n1. BASIC INFORMATION" -ForegroundColor Green
Write-Host "Total records: $($data.Count)"
Write-Host "Columns: $(($data[0].PSObject.Properties.Name) -join ', ')"

# 2. Data overview
Write-Host "`n2. DATA OVERVIEW" -ForegroundColor Green
Get-DataOverview -Data $data

# 3. Detailed analysis
Write-Host "`n3. DETAILED COLUMN ANALYSIS" -ForegroundColor Green
Analyze-DataColumns -Data $data

# 4. Data quality
Write-Host "`n4. DATA QUALITY ASSESSMENT" -ForegroundColor Green
Test-DataQuality -Data $data

# 5. Statistical summary
Write-Host "`n5. STATISTICAL ANALYSIS" -ForegroundColor Green
Get-StatisticalSummary -Data $data

# 6. Memory usage
Write-Host "`n6. SYSTEM INFORMATION" -ForegroundColor Green
$memory = [System.GC]::GetTotalMemory($false) / 1MB
Write-Host "Approximate memory usage: $([math]::Round($memory, 2)) MB"
Write-Host "Data object type: $($data.GetType().Name)"

Write-Host "`nAnalysis complete!" -ForegroundColor Green
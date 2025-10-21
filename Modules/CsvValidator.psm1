function Test-AutopilotCsv {
    <#
    .SYNOPSIS
    Validates Autopilot CSV file structure and content
    
    .PARAMETER FilePath
    Path to CSV file
    #>
    param(
        [Parameter(Mandatory)]
        [string]$FilePath
    )
    
    $errors = @()
    $isValid = $true
    $deviceCount = 0
    
    try {
        # Check file exists
        if (-not (Test-Path $FilePath)) {
            return @{
                IsValid = $false
                Errors = @("File not found: $FilePath")
                DeviceCount = 0
            }
        }
        
        # Import CSV
        $csvData = Import-Csv -Path $FilePath -ErrorAction Stop
        $deviceCount = $csvData.Count
        
        if ($deviceCount -eq 0) {
            $errors += "CSV file is empty"
            $isValid = $false
            return @{
                IsValid = $isValid
                Errors = $errors
                DeviceCount = $deviceCount
            }
        }
        
        # Get column names
        $csvColumns = $csvData[0].PSObject.Properties.Name
        
        # Check for required columns (flexible naming)
        $serialNumberColumn = $csvColumns | Where-Object { 
            $_ -match "Serial|Device Serial" 
        } | Select-Object -First 1
        
        $hardwareHashColumn = $csvColumns | Where-Object { 
            $_ -match "Hardware|Hash" 
        } | Select-Object -First 1
        
        if (-not $serialNumberColumn) {
            $errors += "Missing required column: 'Device Serial Number' or similar"
            $isValid = $false
        }
        
        if (-not $hardwareHashColumn) {
            $errors += "Missing required column: 'Hardware Hash' or similar"
            $isValid = $false
        }
        
        # Validate data integrity
        if ($isValid) {
            $rowNumber = 1
            foreach ($row in $csvData) {
                $rowNumber++
                
                $serialNumber = $row.$serialNumberColumn
                $hardwareHash = $row.$hardwareHashColumn
                
                if ([string]::IsNullOrWhiteSpace($serialNumber)) {
                    $errors += "Row ${rowNumber}: Missing Serial Number"
                    $isValid = $false
                }
                
                if ([string]::IsNullOrWhiteSpace($hardwareHash)) {
                    $errors += "Row ${rowNumber}: Missing Hardware Hash"
                    $isValid = $false
                }
                
                # Basic hardware hash validation (should be long Base64 string)
                if ($hardwareHash -and $hardwareHash.Length -lt 100) {
                    $errors += "Row ${rowNumber}: Hardware Hash appears invalid (too short)"
                }
            }
        }
        
    } catch {
        $errors += "Error reading CSV: $($_.Exception.Message)"
        $isValid = $false
    }
    
    return @{
        IsValid = $isValid
        Errors = $errors
        DeviceCount = $deviceCount
    }
}

Export-ModuleMember -Function Test-AutopilotCsv

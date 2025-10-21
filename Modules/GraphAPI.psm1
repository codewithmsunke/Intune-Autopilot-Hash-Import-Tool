function Import-AutopilotDevices {
    <#
    .SYNOPSIS
    Imports Autopilot devices from CSV to Intune
    
    .PARAMETER CsvPath
    Path to CSV file containing device information
    
    .PARAMETER AccessToken
    Not used - function uses existing MgGraph session
    
    .PARAMETER StatusCallback
    Optional script block for status updates to GUI
    #>
    param(
        [Parameter(Mandatory)]
        [string]$CsvPath,
        
        [Parameter()]
        [string]$AccessToken,
        
        [Parameter()]
        [scriptblock]$StatusCallback
    )
    
    # Helper function to send status updates
    function Send-Status {
        param(
            [string]$Message,
            [string]$Type = "Info"
        )
        if ($StatusCallback) {
            & $StatusCallback $Message $Type
        } else {
            $color = switch ($Type) {
                "Success" { "Green" }
                "Warning" { "Yellow" }
                "Error"   { "Red" }
                default   { "White" }
            }
            Write-Host $Message -ForegroundColor $color
        }
    }
    
    # Verify Graph connection
    $context = Get-MgContext
    if ($null -eq $context) {
        throw "Not connected to Microsoft Graph. Please authenticate first."
    }
    
    # Read CSV file
    $devices = Import-Csv -Path $CsvPath
    
    $graphApiVersion = "beta"
    $Resource = "deviceManagement/importedWindowsAutopilotDeviceIdentities"
    $baseUri = "https://graph.microsoft.com/$graphApiVersion/$Resource"
    
    $successCount = 0
    $failureCount = 0
    [System.Collections.ArrayList]$failedDevices = @()
    
    # Get column names (flexible for different CSV formats)
    $csvColumns = $devices[0].PSObject.Properties.Name
    $serialColumn = $csvColumns | Where-Object { $_ -match "Serial|Device Serial" } | Select-Object -First 1
    $hashColumn = $csvColumns | Where-Object { $_ -match "Hardware|Hash" } | Select-Object -First 1
    
    $currentDevice = 0
    $totalDevices = $devices.Count
    
    foreach ($device in $devices) {
           $currentDevice++
           $serialNumber = $device.$serialColumn
           $hardwareHash = $device.$hashColumn

           # Write device progress line
           Send-Status "$currentDevice of $totalDevices" "Info"

           

           try {
            # Build JSON payload
            $json = @"
{
    "@odata.type": "#microsoft.graph.importedWindowsAutopilotDeviceIdentity",
    "serialNumber": "$serialNumber",
    "productKey": "",
    "hardwareIdentifier": "$hardwareHash",
    "state": {
        "@odata.type": "microsoft.graph.importedWindowsAutopilotDeviceIdentityState",
        "deviceImportStatus": "pending",
        "deviceRegistrationId": "",
        "deviceErrorCode": 0,
        "deviceErrorName": ""
    }
}
"@
            
            # Post import request
            $firstresponse = Invoke-MgGraphRequest -Uri $baseUri -Method Post -Body $json -ContentType "application/json"
            
            $id = $firstresponse.id
            
            # Poll for completion
            if ($id) {
                $pollUri = "$baseUri/$id"
            } else {
                $pollUri = $baseUri
            }
            
            $maxAttempts = 20
            $attempt = 0
            $delaySeconds = 15
            
            $finalStatus = $null
            $finalErrorCode = $null
            $finalErrorName = $null
            do {
                try {
                    

                    $response = Invoke-MgGraphRequest -Uri $pollUri -Method Get
                    # Extract state information dynamically
                    if ($id) {
                        $state = $response.state
                    } else {
                        $state = $response.Value[0].state
                    }
                    $status = $state.deviceImportStatus
                    $errorCode = $state.deviceErrorCode
                    $errorName = $state.deviceErrorName
                    $registrationId = $state.deviceRegistrationId
                    # Show polling progress
                    Send-Status "  Polling attempt $($attempt + 1)/$maxAttempts - Status: $status" "Info"
                    $finalStatus = $status
                    $finalErrorCode = $errorCode
                    $finalErrorName = $errorName
                    if ($status -eq "complete") {
                        $successCount++
                        break
                    } elseif ($status -eq "error") {
                        $failureCount++
                        [void]$failedDevices.Add(@{
                            SerialNumber = $serialNumber
                            ErrorCode = $errorCode
                            ErrorName = $errorName
                            Status = $status
                            RegistrationId = $registrationId
                        })
                        break
                    }
                } catch {
                    # Silently continue polling on transient errors
                }
                $attempt++
                if ($attempt -ge $maxAttempts) {
                    $failureCount++
                    $finalStatus = "timeout"
                    $finalErrorCode = -1
                    $finalErrorName = "Timeout"
                    [void]$failedDevices.Add(@{
                        SerialNumber = $serialNumber
                        ErrorCode = -1
                        ErrorName = "Timeout"
                        Status = "timeout"
                        RegistrationId = $null
                    })
                    break
                }
                Start-Sleep -Seconds $delaySeconds
            } while ($true)

            # After polling, write device import status line
            $desc = if ($finalErrorName) {
                $finalErrorName -creplace '([A-Z])', ' $1'
            } else {
                $finalStatus
            }
            $desc = $desc.Trim()
            Send-Status "$serialNumber - $finalStatus - $finalErrorCode - $desc" $(if ($finalStatus -eq 'complete') { 'Success' } elseif ($finalStatus -eq 'error') { 'Error' } else { 'Warning' })
            
            } catch {
            $ex = $_.Exception
            $failureCount++
            [void]$failedDevices.Add(@{
                SerialNumber = $serialNumber
                ErrorCode = 0
                ErrorName = "Exception"
                Status = "exception"
                RegistrationId = $null
                ExceptionMessage = $ex.Message
            })
        }
    }
    
    # Ensure failureCount matches actual failed devices count
    $failureCount = $failedDevices.Count

    # Overall import summary
    Send-Status "" "Info"
    Send-Status "======================================" "Info"
    Send-Status "Import Summary:" "Info"
    Send-Status "  Total Devices: $totalDevices" "Info"
    Send-Status "  Successfully imported: $successCount" "Success"
    Send-Status "  Failed: $failureCount" $(if ($failureCount -gt 0) { "Error" } else { "Info" })
    Send-Status "======================================" "Info"

    return @{
        SuccessCount = $successCount
        FailureCount = $failureCount
        FailedDevices = $failedDevices
        TotalDevices = $totalDevices
    }
}

Export-ModuleMember -Function Import-AutopilotDevices

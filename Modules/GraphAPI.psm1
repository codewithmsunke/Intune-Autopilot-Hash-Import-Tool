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
    $failedDevices = @()
    
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
        
        try {
            Send-Status "[$currentDevice/$totalDevices] Importing device: $serialNumber" "Info"
            
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
                    
                    Send-Status "  Polling attempt $($attempt + 1)/$maxAttempts - Status: $status" "Info"
                    
                    if ($status -eq "complete") {
                        Send-Status "  ✓ Successfully imported: $serialNumber" "Success"
                        if ($registrationId) {
                            Send-Status "     Registration ID: $registrationId" "Info"
                        }
                        $successCount++
                        break
                        
                    } elseif ($status -eq "error") {
                        # Display error details from Graph API response
                        Send-Status "  ✗ Import Failed: $serialNumber" "Error"
                        
                        if ($errorCode -and $errorCode -ne 0) {
                            Send-Status "     Error Code: $errorCode" "Error"
                        }
                        
                        if ($errorName) {
                            # Convert camelCase error name to readable format
                            $readableError = $errorName -creplace '([A-Z])', ' $1'
                            $readableError = $readableError.Trim()
                            Send-Status "     Error: $readableError" "Error"
                        }
                        
                        # Provide context-specific guidance based on error
                        if ($errorName -like "*AlreadyAssigned*" -or $errorCode -eq 806) {
                            Send-Status "     ℹ Device is already registered in Autopilot" "Warning"
                            Send-Status "     Solution: Delete the device from Intune Autopilot portal and try again" "Warning"
                        } elseif ($errorName -like "*Invalid*" -or $errorCode -in @(807, 808)) {
                            Send-Status "     ℹ Hardware hash or serial number format is invalid" "Warning"
                            Send-Status "     Solution: Re-run Get-WindowsAutoPilotInfo.ps1 to generate a new CSV" "Warning"
                        } elseif ($errorCode -gt 0) {
                            Send-Status "     ℹ Check the device hardware hash and serial number" "Warning"
                        }
                        
                        $failureCount++
                        # Store detailed error information
                        $failedDevices += @{
                            SerialNumber = $serialNumber
                            ErrorCode = $errorCode
                            ErrorName = $errorName
                            Status = $status
                            RegistrationId = $registrationId
                        }
                        break
                    }
                    
                } catch {
                    Send-Status "  ⚠ Polling error: $($_.Exception.Message)" "Warning"
                }
                
                $attempt++
                if ($attempt -ge $maxAttempts) {
                    Send-Status "  ✗ Timeout: $serialNumber - Max polling attempts reached" "Error"
                    Send-Status "     ℹ Import may still be processing in the background" "Warning"
                    $failureCount++
                    $failedDevices += @{
                        SerialNumber = $serialNumber
                        ErrorCode = -1
                        ErrorName = "Timeout"
                        Status = "timeout"
                        RegistrationId = $null
                    }
                    break
                }
                
                Start-Sleep -Seconds $delaySeconds
                
            } while ($true)
            
        } catch {
            $ex = $_.Exception
            Send-Status "  ✗ Failed to import: $serialNumber" "Error"
            Send-Status "     Exception: $($ex.Message)" "Error"
            
            # Try to extract Graph API error details from response
            if ($ex.Response) {
                try {
                    $errorResponse = $ex.Response.GetResponseStream()
                    $reader = New-Object System.IO.StreamReader($errorResponse)
                    $reader.BaseStream.Position = 0
                    $reader.DiscardBufferedData()
                    $responseBody = $reader.ReadToEnd()
                    
                    # Parse JSON error if available
                    try {
                        $errorJson = $responseBody | ConvertFrom-Json
                        if ($errorJson.error) {
                            Send-Status "     API Error: $($errorJson.error.message)" "Error"
                            if ($errorJson.error.code) {
                                Send-Status "     API Code: $($errorJson.error.code)" "Error"
                            }
                        }
                    } catch {
                        # If not JSON, show raw response
                        Send-Status "     Details: $responseBody" "Error"
                    }
                } catch {
                    # Ignore if can't read error details
                }
            }
            
            $failureCount++
            $failedDevices += @{
                SerialNumber = $serialNumber
                ErrorCode = 0
                ErrorName = "Exception"
                Status = "exception"
                RegistrationId = $null
                ExceptionMessage = $ex.Message
            }
        }
    }
    
    # Return detailed results with summary
    Send-Status "" "Info"
    Send-Status "======================================" "Info"
    Send-Status "Import Summary:" "Info"
    Send-Status "  Total Devices: $totalDevices" "Info"
    Send-Status "  Successful: $successCount" "Success"
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

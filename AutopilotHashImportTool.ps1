<#
.SYNOPSIS
    Autopilot Hash Import Tool - GUI application for importing device hashes to Intune Autopilot

.DESCRIPTION
    This tool provides a user-friendly interface to:
    1. Authenticate to Microsoft Graph
    2. Select and validate CSV files containing Autopilot device hashes
    3. Import devices to Intune Autopilot with real-time progress tracking

.NOTES
    Requirements:
    - PowerShell 5.1 or later
    - Microsoft.Graph.Authentication module (auto-installed)
    - Valid Azure AD app registration
    - Appropriate Intune permissions
#>

# Load required assemblies
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# Import modules
$modulePath = "$PSScriptRoot\Modules"
Import-Module "$modulePath\Authentication.psm1" -Force
Import-Module "$modulePath\GraphAPI.psm1" -Force
Import-Module "$modulePath\CsvValidator.psm1" -Force

# Load configuration
$configPath = "$PSScriptRoot\Config\AppConfig.json"
if (-not (Test-Path $configPath)) {
    [System.Windows.MessageBox]::Show(
        "Configuration file not found: $configPath`n`nPlease create AppConfig.json with ClientId and TenantId.",
        "Configuration Error",
        "OK",
        "Error"
    )
    exit
}

$config = Get-Content $configPath | ConvertFrom-Json

# Validate config
if ([string]::IsNullOrWhiteSpace($config.ClientId) -or $config.ClientId -like "*PASTE*") {
    [System.Windows.MessageBox]::Show(
        "Please update AppConfig.json with your Azure AD App Registration details:`n`n" +
        "- ClientId: Your Application (client) ID`n" +
        "- TenantId: Your Directory (tenant) ID",
        "Configuration Required",
        "OK",
        "Warning"
    )
}

# Define XAML for UI
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Autopilot Hash Import Tool" Height="650" Width="850"
        WindowStartupLocation="CenterScreen" ResizeMode="CanResize"
        Background="#F5F5F5">
    <Window.Resources>
        <Style TargetType="GroupBox">
            <Setter Property="Margin" Value="0,0,0,15"/>
            <Setter Property="Padding" Value="15"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="BorderBrush" Value="#DDDDDD"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>
        <Style TargetType="Button">
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="10,5"/>
        </Style>
    </Window.Resources>
    
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Border Grid.Row="0" Background="#0078D4" Padding="15" Margin="0,0,0,20" CornerRadius="5">
            <StackPanel>
                <TextBlock Text="Autopilot Hash Import Tool" FontSize="20" 
                           FontWeight="Bold" Foreground="White"/>
                <TextBlock Text="Import Windows Autopilot device hashes to Microsoft Intune" 
                           FontSize="12" Foreground="#E0E0E0" Margin="0,5,0,0"/>
            </StackPanel>
        </Border>

        <!-- Main Content -->
        <StackPanel Grid.Row="1">
            <!-- 1. Authentication Section -->
            <GroupBox Header="1. Authentication">
                <StackPanel>
                    <Button Name="btnAuthenticate" Content="Sign In to Microsoft Graph" 
                            Height="45" FontSize="14" Background="#0078D4" 
                            Foreground="White" Margin="0,0,0,10"/>
                    <Border Background="#F9F9F9" Padding="10" CornerRadius="3">
                        <TextBlock Name="lblUserInfo" Text="Status: Not signed in" 
                                   FontSize="12" Foreground="#666"/>
                    </Border>
                </StackPanel>
            </GroupBox>

            <!-- 2. CSV Selection Section -->
            <GroupBox Header="2. CSV File Selection">
                <StackPanel>
                    <Grid Margin="0,0,0,10">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBox Name="txtFilePath" Grid.Column="0" Height="40" 
                                 VerticalContentAlignment="Center" IsReadOnly="True" 
                                 Margin="0,0,10,0" FontSize="12" Background="#F9F9F9"
                                 Padding="10"/>
                        <Button Name="btnBrowse" Grid.Column="1" Content="Browse CSV" 
                                Width="130" Height="40" IsEnabled="False"
                                Background="#6C757D" Foreground="White"/>
                    </Grid>
                    <Border Background="#F9F9F9" Padding="10" CornerRadius="3" Margin="0,0,0,10">
                        <TextBlock Name="lblFileInfo" Text="No file selected" 
                                   FontSize="11" Foreground="#666"/>
                    </Border>
                    <Button Name="btnImport" Content="Import to Autopilot" 
                            Height="45" FontSize="14" IsEnabled="False" 
                            Background="#107C10" Foreground="White"/>
                </StackPanel>
            </GroupBox>
        </StackPanel>

        <!-- 3. Status Section -->
        <GroupBox Grid.Row="2" Header="3. Import Status and Progress">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                
                <!-- Progress Bar -->
                <StackPanel Grid.Row="0" Margin="0,0,0,10">
                    <ProgressBar Name="progressBar" Height="35" Visibility="Collapsed" 
                                 IsIndeterminate="True" Foreground="#0078D4"/>
                    <TextBlock Name="lblProgress" Text="" FontSize="12" 
                               Margin="0,8,0,0" FontWeight="SemiBold" 
                               HorizontalAlignment="Center" Foreground="#0078D4"/>
                </StackPanel>
                
                <!-- Status Log -->
                <Border Grid.Row="1" BorderBrush="#DDDDDD" BorderThickness="1" 
                        Background="#FAFAFA" CornerRadius="3">
                    <ScrollViewer Name="scrollViewer" VerticalScrollBarVisibility="Auto" 
                                  Padding="10">
                        <TextBlock Name="txtStatus" TextWrapping="Wrap" 
                                   FontFamily="Consolas" FontSize="11"/>
                    </ScrollViewer>
                </Border>
            </Grid>
        </GroupBox>

        <!-- Footer -->
        <Border Grid.Row="3" Background="White" Padding="10" CornerRadius="3" 
                BorderBrush="#DDDDDD" BorderThickness="1" Margin="0,10,0,0">
            <Grid>
                <TextBlock Text="Version 1.0.0 | Powered by Microsoft Graph API" 
                           FontSize="10" Foreground="#999" HorizontalAlignment="Left"/>
                <TextBlock Text="Autopilot Hash Import Tool" FontSize="10" Foreground="#0078D4" 
                           HorizontalAlignment="Right" FontWeight="SemiBold"/>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# Load XAML
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

# Get UI controls
$btnAuthenticate = $window.FindName("btnAuthenticate")
$btnBrowse = $window.FindName("btnBrowse")
$btnImport = $window.FindName("btnImport")
$txtFilePath = $window.FindName("txtFilePath")
$txtStatus = $window.FindName("txtStatus")
$lblUserInfo = $window.FindName("lblUserInfo")
$lblFileInfo = $window.FindName("lblFileInfo")
$lblProgress = $window.FindName("lblProgress")
$progressBar = $window.FindName("progressBar")
$scrollViewer = $window.FindName("scrollViewer")

# Global state variables
$script:isAuthenticated = $false
$script:selectedCsvPath = $null
$script:validationResult = $null

# Synchronized hashtable for runspace communication
$syncHash = [hashtable]::Synchronized(@{})
$syncHash.StatusQueue = [System.Collections.Queue]::Synchronized((New-Object System.Collections.Queue))
$syncHash.IsImportRunning = $false

# Helper function: Write to status log
function Write-StatusLog {
    param(
        [string]$Message,
        [string]$Type = "Info"
    )
    
    $timestamp = Get-Date -Format "HH:mm:ss"
    $color = switch ($Type) {
        "Success" { "#107C10" }
        "Warning" { "#F7630C" }
        "Error"   { "#D13438" }
        default   { "#333333" }
    }
    
    $icon = switch ($Type) {
        "Success" { "✓" }
        "Warning" { "⚠" }
        "Error"   { "✗" }
        default   { "•" }
    }
    
    $run = New-Object System.Windows.Documents.Run
    $run.Text = "[$timestamp] $icon $Message`n"
    $run.Foreground = $color
    $txtStatus.Inlines.Add($run)
    
    $scrollViewer.ScrollToEnd()
    $window.Dispatcher.Invoke([Action]{}, "Render")
}

# Status update timer - polls queue and updates GUI
$statusTimer = New-Object System.Windows.Threading.DispatcherTimer
$statusTimer.Interval = [TimeSpan]::FromMilliseconds(100)
$statusTimer.Add_Tick({
    # Process all queued status messages
    while ($syncHash.StatusQueue.Count -gt 0) {
        try {
            $statusItem = $syncHash.StatusQueue.Dequeue()
            Write-StatusLog -Message $statusItem.Message -Type $statusItem.Type
        } catch {
            # Ignore dequeue errors
        }
    }
})
$statusTimer.Start()

# Helper function: Update progress bar
function Update-Progress {
    param(
        [string]$Message = "",
        [bool]$Show = $false
    )
    
    if ($Show) {
        $progressBar.Visibility = "Visible"
        $lblProgress.Text = $Message
    } else {
        $progressBar.Visibility = "Collapsed"
        $lblProgress.Text = ""
    }
    $window.Dispatcher.Invoke([Action]{}, "Render")
}

# Event: Authenticate Button Click
$btnAuthenticate.Add_Click({
    if ($script:isAuthenticated) {
        # Sign Out
        try {
            Disconnect-MgGraph -ErrorAction SilentlyContinue
            $script:isAuthenticated = $false
            $lblUserInfo.Text = "Status: Signed out"
            $lblUserInfo.Foreground = "#D13438"
            $btnAuthenticate.Content = "Sign In to Microsoft Graph"
            $btnAuthenticate.Background = "#0078D4"
            $btnBrowse.IsEnabled = $false
            $btnImport.IsEnabled = $false
            Write-StatusLog "User signed out" "Warning"
        } catch {
            Write-StatusLog "Sign out error: $($_.Exception.Message)" "Error"
        }
    } else {
        # Sign In
        try {
            Write-StatusLog "Initiating authentication..." "Info"
            Update-Progress "Authenticating to Microsoft Graph..." $true
            $btnAuthenticate.IsEnabled = $false
            
            $authResult = Connect-GraphInteractive `
                -ClientId $config.ClientId `
                -TenantId $config.TenantId
            
            $script:isAuthenticated = $true
            $lblUserInfo.Text = "Status: [OK] Signed in as $($authResult.Account.Username)"
            $lblUserInfo.Foreground = "#107C10"
            $btnAuthenticate.Content = "Sign Out"
            $btnAuthenticate.Background = "#F7630C"
            $btnBrowse.IsEnabled = $true
            
            Write-StatusLog "Successfully authenticated as $($authResult.Account.Username)" "Success"
            Write-StatusLog "You can now select a CSV file to import" "Info"
            
        } catch {
            Write-StatusLog "Authentication failed: $($_.Exception.Message)" "Error"
            [System.Windows.MessageBox]::Show(
                "Authentication failed!`n`nError: $($_.Exception.Message)`n`n" +
                "Please ensure:`n" +
                "• You have internet connectivity`n" +
                "• App registration is configured correctly`n" +
                "• You have appropriate Intune permissions",
                "Authentication Error",
                "OK",
                "Error"
            )
        } finally {
            Update-Progress "" $false
            $btnAuthenticate.IsEnabled = $true
        }
    }
})

# Event: Browse Button Click
$btnBrowse.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $openFileDialog.Title = "Select Autopilot Hash CSV File"
    $openFileDialog.Multiselect = $false
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:selectedCsvPath = $openFileDialog.FileName
        $txtFilePath.Text = $script:selectedCsvPath
        
        $fileName = [System.IO.Path]::GetFileName($script:selectedCsvPath)
        Write-StatusLog "CSV file selected: $fileName" "Info"
        
        # Validate CSV
        Write-StatusLog "Validating CSV structure..." "Info"
        Update-Progress "Validating CSV file..." $true
        
        try {
            $script:validationResult = Test-AutopilotCsv -FilePath $script:selectedCsvPath
            
            if ($script:validationResult.IsValid) {
                $lblFileInfo.Text = "✓ Valid CSV | $($script:validationResult.DeviceCount) device(s) ready to import"
                $lblFileInfo.Foreground = "#107C10"
                $btnImport.IsEnabled = $true
                Write-StatusLog "Validation passed: $($script:validationResult.DeviceCount) device(s) found" "Success"
                Write-StatusLog "Click 'Import to Autopilot' to begin import process" "Info"
            } else {
                $lblFileInfo.Text = "✗ Invalid CSV | Please check file format"
                $lblFileInfo.Foreground = "#D13438"
                $btnImport.IsEnabled = $false
                Write-StatusLog "CSV validation failed" "Error"
                foreach ($validationError in $script:validationResult.Errors) {
                    Write-StatusLog "  $validationError" "Error"
                }
                
                [System.Windows.MessageBox]::Show(
                    "CSV validation failed:`n`n$($script:validationResult.Errors -join "`n")`n`n" +
                    "Required columns: Device Serial Number, Hardware Hash",
                    "Validation Error",
                    "OK",
                    "Warning"
                )
            }
        } catch {
            $lblFileInfo.Text = "✗ Error reading CSV file"
            $lblFileInfo.Foreground = "#D13438"
            Write-StatusLog "Error validating CSV: $($_.Exception.Message)" "Error"
        } finally {
            Update-Progress "" $false
        }
    }
})

# Event: Import Button Click
$btnImport.Add_Click({
    # Prevent multiple simultaneous imports
    if ($syncHash.IsImportRunning) {
        [System.Windows.MessageBox]::Show(
            "An import is already in progress. Please wait for it to complete.",
            "Import In Progress",
            "OK",
            "Warning"
        )
        return
    }
    
    $confirmResult = [System.Windows.MessageBox]::Show(
        "Import $($script:validationResult.DeviceCount) device(s) to Autopilot?`n`n" +
        "This process may take several minutes depending on the number of devices.`n`n" +
        "The UI will remain responsive - you can minimize or view other windows.",
        "Confirm Import",
        "YesNo",
        "Question"
    )
    
    if ($confirmResult -eq "Yes") {
        # Disable controls during import
        $btnImport.IsEnabled = $false
        $btnBrowse.IsEnabled = $false
        $btnAuthenticate.IsEnabled = $false
        $syncHash.IsImportRunning = $true
        
        # Show progress
        Update-Progress "Importing devices to Autopilot (this may take several minutes)..." $true
        
        Write-StatusLog "======================================" "Info"
        Write-StatusLog "Starting import process for $($script:validationResult.DeviceCount) device(s)" "Info"
        Write-StatusLog "======================================" "Info"
        
        # Create runspace for background processing
        $runspace = [runspacefactory]::CreateRunspace()
        $runspace.ApartmentState = "STA"
        $runspace.ThreadOptions = "ReuseThread"
        $runspace.Open()
        
        # Pass variables to runspace
        $runspace.SessionStateProxy.SetVariable("syncHash", $syncHash)
        $runspace.SessionStateProxy.SetVariable("csvPath", $script:selectedCsvPath)
        $runspace.SessionStateProxy.SetVariable("modulePath", "$PSScriptRoot\Modules")
        
        # Create PowerShell instance with script block
        $ps = [powershell]::Create()
        $ps.Runspace = $runspace
        
        [void]$ps.AddScript({
            param($syncHash, $csvPath, $modulePath)
            
            # Import modules in runspace context
            Import-Module "$modulePath\GraphAPI.psm1" -Force
            
            # Define status callback to queue messages for GUI
            $statusCallback = {
                param($Message, $Type = "Info")
                try {
                    $syncHash.StatusQueue.Enqueue(@{
                        Message = $Message
                        Type = $Type
                        Timestamp = Get-Date
                    })
                } catch {
                    # Ignore queue errors
                }
            }
            
            # Execute import with callback
            try {
                $result = Import-AutopilotDevices `
                    -CsvPath $csvPath `
                    -AccessToken $null `
                    -StatusCallback $statusCallback
                
                # Store result in sync hash
                $syncHash.ImportResult = $result
                $syncHash.ImportSuccess = $true
                
            } catch {
                $syncHash.ImportError = $_.Exception.Message
                $syncHash.ImportSuccess = $false
                
                # Queue error message
                & $statusCallback "EXCEPTION: $($_.Exception.Message)" "Error"
            } finally {
                $syncHash.ImportComplete = $true
            }
            
        }).AddArgument($syncHash).AddArgument($script:selectedCsvPath).AddArgument("$PSScriptRoot\Modules")
        
        # Start async execution
        $asyncResult = $ps.BeginInvoke()
        
        # Create completion monitoring timer
        $completionTimer = New-Object System.Windows.Threading.DispatcherTimer
        $completionTimer.Interval = [TimeSpan]::FromMilliseconds(500)
        $completionTimer.Tag = @{
            PowerShell = $ps
            Runspace = $runspace
            AsyncResult = $asyncResult
        }
        
        $completionTimer.Add_Tick({
            param($sender, $e)
            
            # Check if import is complete
            if ($syncHash.ImportComplete) {
                $sender.Stop()
                
                # Hide progress
                Update-Progress "" $false
            
                # Process results
                if ($syncHash.ImportSuccess -and $syncHash.ImportResult) {
                    $importResult = $syncHash.ImportResult
                    
                    # Display detailed results
                    Write-StatusLog "" "Info"
                    Write-StatusLog "======================================" "Info"
                    Write-StatusLog "Import process completed!" "Success"
                    Write-StatusLog "  Total devices: $($importResult.TotalDevices)" "Info"
                    Write-StatusLog "  Successfully imported: $($importResult.SuccessCount)" "Success"
                    Write-StatusLog "  Failed: $($importResult.FailureCount)" $(if ($importResult.FailureCount -gt 0) { "Error" } else { "Info" })
                    Write-StatusLog "======================================" "Info"
                    
                    if ($importResult.FailureCount -gt 0) {
                        Write-StatusLog "" "Info"
                        Write-StatusLog "Failed Devices - Detailed Error Information:" "Error"
                        Write-StatusLog "======================================" "Info"
                        
                        foreach ($device in $importResult.FailedDevices) {
                            Write-StatusLog "" "Info"
                            Write-StatusLog "Device: $($device.SerialNumber)" "Error"
                            
                            if ($device.ErrorCode -and $device.ErrorCode -ne 0 -and $device.ErrorCode -ne -1) {
                                Write-StatusLog "  Error Code: $($device.ErrorCode)" "Error"
                            }
                            
                            if ($device.ErrorName) {
                                $readableError = $device.ErrorName -creplace '([A-Z])', ' $1'
                                $readableError = $readableError.Trim()
                                Write-StatusLog "  Error Name: $readableError" "Error"
                            }
                            
                            if ($device.Status) {
                                Write-StatusLog "  Status: $($device.Status)" $(if ($device.Status -eq "timeout") { "Warning" } else { "Error" })
                            }
                            
                            if ($device.ErrorName -like "*AlreadyAssigned*" -or $device.ErrorCode -eq 806) {
                                Write-StatusLog "  ℹ This device is already registered in Autopilot" "Warning"
                                Write-StatusLog "  Solution: Go to Intune > Devices > Windows > Enrollment > Devices" "Warning"
                                Write-StatusLog "            Search for serial number and delete, then retry" "Warning"
                            } elseif ($device.ErrorName -like "*Invalid*" -or $device.ErrorCode -in @(807, 808)) {
                                Write-StatusLog "  ℹ Hardware hash or serial number format is invalid" "Warning"
                                Write-StatusLog "  Solution: Re-run Get-WindowsAutoPilotInfo.ps1 to generate new CSV" "Warning"
                            } elseif ($device.ErrorCode -eq -1) {
                                Write-StatusLog "  ℹ Import may still be processing in the background" "Warning"
                                Write-StatusLog "  Solution: Check Intune portal in a few minutes" "Warning"
                            } elseif ($device.ExceptionMessage) {
                                Write-StatusLog "  Exception: $($device.ExceptionMessage)" "Error"
                            }
                            
                            if ($device.RegistrationId) {
                                Write-StatusLog "  Registration ID: $($device.RegistrationId)" "Info"
                            }
                        }
                        Write-StatusLog "======================================" "Info"
                    }
                    
                    # Create summary message
                    $summary = "Import Process Completed!`n`n"
                    $summary += "Total Devices: $($importResult.TotalDevices)`n"
                    $summary += "✓ Successfully imported: $($importResult.SuccessCount)`n"
                    
                    if ($importResult.FailureCount -gt 0) {
                        $alreadyAssigned = ($importResult.FailedDevices | Where-Object { 
                            $_.ErrorName -like "*AlreadyAssigned*" -or $_.ErrorCode -eq 806 
                        }).Count
                        
                        $timeouts = ($importResult.FailedDevices | Where-Object { 
                            $_.ErrorCode -eq -1 
                        }).Count
                        
                        $otherErrors = $importResult.FailureCount - $alreadyAssigned - $timeouts
                        
                        $summary += "`nFailed Imports:`n"
                        if ($alreadyAssigned -gt 0) {
                            $summary += "  ⚠ $alreadyAssigned already registered in Autopilot`n"
                        }
                        if ($timeouts -gt 0) {
                            $summary += "  ⏱ $timeouts timed out (may still be processing)`n"
                        }
                        if ($otherErrors -gt 0) {
                            $summary += "  ✗ $otherErrors failed with errors`n"
                        }
                        
                        $summary += "`nSee status log above for detailed error information and solutions."
                    } else {
                        $summary += "`nAll devices imported successfully!"
                    }
                    
                    $messageType = if ($importResult.FailureCount -eq 0) { 
                        "Information" 
                    } elseif ($importResult.SuccessCount -gt 0) { 
                        "Warning" 
                    } else { 
                        "Error" 
                    }
                    
                    [System.Windows.MessageBox]::Show(
                        $summary,
                        "Import Complete",
                        "OK",
                        $messageType
                    )
                    
                } elseif ($syncHash.ImportError) {
                    Write-StatusLog "Import process failed: $($syncHash.ImportError)" "Error"
                    [System.Windows.MessageBox]::Show(
                        "Import failed!`n`n$($syncHash.ImportError)",
                        "Import Error",
                        "OK",
                        "Error"
                    )
                }
                
                # Cleanup
                $timerTag = $sender.Tag
                try {
                    $timerTag.AsyncResult.AsyncWaitHandle.Close()
                    $timerTag.PowerShell.Dispose()
                    $timerTag.Runspace.Close()
                    $timerTag.Runspace.Dispose()
                } catch {
                    # Ignore cleanup errors
                }
                
                # Reset state
                $syncHash.ImportComplete = $false
                $syncHash.ImportSuccess = $false
                $syncHash.ImportResult = $null
                $syncHash.ImportError = $null
                $syncHash.IsImportRunning = $false
                
                # Re-enable controls
                $btnImport.IsEnabled = $true
                $btnBrowse.IsEnabled = $true
                $btnAuthenticate.IsEnabled = $true
            }
        })
        
        $completionTimer.Start()
    }
})

# Initialize application
Write-StatusLog "Autopilot Hash Import Tool initialized" "Info"
Write-StatusLog "Step 1: Click 'Sign In' to authenticate with Microsoft Graph" "Info"
Write-StatusLog "" "Info"
Write-StatusLog "App Configuration:" "Info"
Write-StatusLog "  Client ID: $($config.ClientId)" "Info"
Write-StatusLog "  Tenant ID: $($config.TenantId)" "Info"

# Show window
$null = $window.ShowDialog()

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
        WindowState="Maximized"
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
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Left">
            <Button Name="btnImport" Content="Import to Autopilot" 
                Height="45" FontSize="14" IsEnabled="False" 
                Background="#107C10" Foreground="White" Width="170" Margin="0,0,10,0"/>
            <Button Name="btnCopySummary" Content="Copy Summary" 
                Height="45" FontSize="14" IsEnabled="False" 
                Background="#0078D4" Foreground="White" Width="130"/>
            </StackPanel>
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
                    <ProgressBar Name="progressBar" Height="20" Visibility="Collapsed" 
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
                        <RichTextBox Name="txtStatus" IsReadOnly="True" IsDocumentEnabled="True"
                                     VerticalScrollBarVisibility="Hidden" BorderThickness="0"
                                     Background="Transparent" FontFamily="Consolas" FontSize="11"
                                     Padding="0"/>
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
$btnCopySummary = $window.FindName("btnCopySummary")

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
    
    $paragraph = $null
    if (-not $script:StatusParagraph) {
        $doc = $txtStatus.Document
        $script:StatusParagraph = New-Object System.Windows.Documents.Paragraph
        $doc.Blocks.Clear()
        $doc.Blocks.Add($script:StatusParagraph)
    }

    $run = New-Object System.Windows.Documents.Run
    $run.Text = "[$timestamp] $icon $Message`n"
    $run.Foreground = $color
    $script:StatusParagraph.Inlines.Add($run)

    # Ensure scroll to end of ScrollViewer that wraps the RichTextBox
    $scrollViewer.ScrollToEnd()
    $window.Dispatcher.Invoke([Action]{}, "Render")

    # Also append exact UI output to audit log on disk (create folder if needed)
    try {
        $logDir = 'C:\wbg\logs'
        if (-not (Test-Path -Path $logDir)) {
            New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        }

        $logPath = Join-Path $logDir 'HashImport.log'
        $logTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $uiIcon = $icon
        $uiLine = "[$logTimestamp] $uiIcon $Message"

        # Rotate log if it exceeds or equals 2 MB
        if (Test-Path -Path $logPath) {
            try {
                $fileInfo = Get-Item -Path $logPath
                if ($fileInfo.Length -ge 2MB) {
                    $archiveName = "HashImport_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss")
                    $archivePath = Join-Path $logDir $archiveName
                    Rename-Item -Path $logPath -NewName $archiveName -Force
                }
            } catch {
                # If rotation fails, continue and append to existing log
            }
        }

        Add-Content -Path $logPath -Value $uiLine -Encoding UTF8
    } catch {
        # If logging fails, don't disrupt the UI
    }
}

# Helper: Safely set text on various WPF controls (TextBlock, TextBox, etc.)
function Set-ControlText {
    param(
        [Parameter(Mandatory=$true)][object]$Control,
        [Parameter(Mandatory=$true)][string]$Text
    )

    if (-not $Control) { return }

    try {
        # Try direct property first
        $Control.Text = $Text
        return
    } catch {}

    try {
        # Try common WPF Text dependency properties
        $Control.SetValue([System.Windows.Controls.TextBlock]::TextProperty, $Text)
        return
    } catch {}

    try {
        $Control.SetValue([System.Windows.Controls.TextBox]::TextProperty, $Text)
        return
    } catch {}

    try {
        $Control.SetValue([System.Windows.Documents.Run]::TextProperty, $Text)
        return
    } catch {}

    # Last resort: try ToString on control and ignore
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
        $lblProgress.SetValue([System.Windows.Controls.TextBlock]::TextProperty, $Message)
    } else {
        $progressBar.Visibility = "Collapsed"
        $lblProgress.SetValue([System.Windows.Controls.TextBlock]::TextProperty, "")
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
            Set-ControlText -Control $lblUserInfo -Text "Status: Signed out"
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
            Set-ControlText -Control $lblUserInfo -Text "Status: [OK] Signed in as $($authResult.Account.Username)"
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
    Set-ControlText -Control $txtFilePath -Text $script:selectedCsvPath
        
        $fileName = [System.IO.Path]::GetFileName($script:selectedCsvPath)
        Write-StatusLog "CSV file selected: $fileName" "Info"
        
        # Validate CSV
        Write-StatusLog "Validating CSV structure..." "Info"
        Update-Progress "Validating CSV file..." $true
        
        try {
            $script:validationResult = Test-AutopilotCsv -FilePath $script:selectedCsvPath
            
            if ($script:validationResult.IsValid) {
                Set-ControlText -Control $lblFileInfo -Text "✓ Valid CSV | $($script:validationResult.DeviceCount) device(s) ready to import"
                $lblFileInfo.Foreground = "#107C10"
                $btnImport.IsEnabled = $true
                Write-StatusLog "Validation passed: $($script:validationResult.DeviceCount) device(s) found" "Success"
                Write-StatusLog "Click 'Import to Autopilot' to begin import process" "Info"
            } else {
                Set-ControlText -Control $lblFileInfo -Text "✗ Invalid CSV | Please check file format"
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
            Set-ControlText -Control $lblFileInfo -Text "✗ Error reading CSV file"
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

        # RESET PER-IMPORT STATE: prevent aggregation from previous runs
        $syncHash.ImportComplete = $false
        $syncHash.ImportResult = $null
        # Clear status queue safely
        while ($syncHash.StatusQueue.Count -gt 0) {
            try { [void]$syncHash.StatusQueue.Dequeue() } catch { break }
        }

        # Show progress
        Update-Progress "Importing devices to Autopilot (this may take several minutes)..." $true
        Write-StatusLog "======================================" "Info"
        Write-StatusLog "Starting import: $($script:validationResult.DeviceCount) device(s)" "Info"
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
                    
                    # Create summary message
                    $summary = "Import Process Completed!`n`n"
                    $summary += "Total Devices: $($importResult.TotalDevices)`n"
                    $summary += "✓ Successfully imported: $($importResult.SuccessCount)`n"
                    
                    if ($importResult.FailureCount -gt 0) {
                        # Force array to ensure .Count works correctly (single hashtable returns key count, not 1)
                        $alreadyAssignedList = @($importResult.FailedDevices | Where-Object { 
                            $_.ErrorName -like "*AlreadyAssigned*" -or $_.ErrorCode -eq 806 
                        })
                        $alreadyAssigned = $alreadyAssignedList.Count
                        
                        $timeoutsList = @($importResult.FailedDevices | Where-Object { 
                            $_.ErrorCode -eq -1 
                        })
                        $timeouts = $timeoutsList.Count
                        
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

                    # Enable Copy Summary button and store last summary
                    try { $window.FindName('btnCopySummary').IsEnabled = $true } catch {}
                    $script:LastImportSummary = $summary
                    
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

# Copy Summary button handler
$btnCopySummary.Add_Click({
    try {
        if ($script:LastImportSummary) {
            [Windows.Clipboard]::SetText($script:LastImportSummary)
            Write-StatusLog "Summary copied to clipboard" "Success"
        } else {
            Write-StatusLog "No summary available to copy" "Warning"
        }
    } catch {
        Write-StatusLog "Failed to copy summary to clipboard: $($_.Exception.Message)" "Error"
    }
})

# Initialize application
Write-StatusLog "Autopilot Hash Import Tool initialized" "Info"
Write-StatusLog "Step 1: Click 'Sign In' to authenticate with Microsoft Graph" "Info"

# Show window
$null = $window.ShowDialog()

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

# Set window icon from embedded Base64 PNG
$iconBase64 = @'
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
'@

$iconBytes = [Convert]::FromBase64String($iconBase64)
$ms = New-Object System.IO.MemoryStream(,$iconBytes)
$bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
$bitmap.BeginInit()
$bitmap.StreamSource = $ms
$bitmap.CacheOption = 'OnLoad'
$bitmap.DecodePixelWidth = 128  # Set icon width to 64px for better visibility
$bitmap.DecodePixelHeight = 128 # Set icon height to 64px
$bitmap.EndInit()
$bitmap.Freeze()
$window.Icon = $bitmap

# Show window
$null = $window.ShowDialog()

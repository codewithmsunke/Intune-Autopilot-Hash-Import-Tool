function Connect-GraphInteractive {
    <#
    .SYNOPSIS
    Authenticates to Microsoft Graph using delegated permissions
    
    .PARAMETER ClientId
    Application (client) ID from app registration
    
    .PARAMETER TenantId
    Directory (tenant) ID
    #>
    param(
        [Parameter(Mandatory)]
        [string]$ClientId,
        
        [Parameter(Mandatory)]
        [string]$TenantId
    )
    
    # Check and install Microsoft.Graph.Authentication module
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        Write-Host "Installing Microsoft.Graph.Authentication module..." -ForegroundColor Yellow
        Install-Module -Name Microsoft.Graph.Authentication -Force -Scope CurrentUser -AllowClobber
    }
    
    Import-Module Microsoft.Graph.Authentication
    
    try {
        # Disconnect existing session
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        
        # Interactive authentication with required scopes
        Connect-MgGraph -ClientId $ClientId -TenantId $TenantId `
            -Scopes "DeviceManagementServiceConfig.ReadWrite.All","User.Read" `
            -NoWelcome -ErrorAction Stop
        
        # Get context
        $context = Get-MgContext
        
        if ($null -eq $context) {
            throw "Failed to establish Graph context. You may not have access to this application."
        }
        
        return @{
            AccessToken = $context.AccessToken
            Account = @{
                Username = $context.Account
            }
        }
        
    } catch {
        $errorMessage = $_.Exception.Message
        
        # Handle specific error cases and re-throw with user-friendly messages
        if ($errorMessage -match "AADSTS50105") {
            throw "Access Denied: You are not assigned to this application. Please contact your administrator."
        } elseif ($errorMessage -match "AADSTS65001|AADSTS65004") {
            throw "Consent Required: Admin consent is required for this application. Please contact your administrator."
        } elseif ($errorMessage -match "Cancelled|User cancel") {
            throw "Authentication cancelled by user."
        } else {
            throw $errorMessage
        }
    }
}

Export-ModuleMember -Function Connect-GraphInteractive

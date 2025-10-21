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
            -NoWelcome
        
        # Get context
        $context = Get-MgContext
        
        if ($null -eq $context) {
            throw "Failed to establish Graph context"
        }
        
        return @{
            AccessToken = $context.AccessToken
            Account = @{
                Username = $context.Account
            }
        }
        
    } catch {
        throw "Authentication failed: $($_.Exception.Message)"
    }
}

Export-ModuleMember -Function Connect-GraphInteractive

# Azure App Registration Setup Guide

This guide provides detailed step-by-step instructions for setting up Azure AD App Registration for the Autopilot Hash Import Tool.

## Prerequisites

- Azure AD tenant with Intune licenses
- Global Administrator or Application Administrator role
- Access to Azure Portal

---

## Part 1: Register the Application

### Step 1: Navigate to Azure Portal

1. Open your web browser and go to [https://portal.azure.com](https://portal.azure.com)
2. Sign in with your administrator account

### Step 2: Access App Registrations

1. In the Azure Portal, click on **Microsoft Entra ID** (formerly Azure Active Directory)
   - You can find it in the left navigation menu or search for it in the top search bar
2. In the left menu, under **Manage**, click **App registrations**
3. Click **+ New registration** at the top

### Step 3: Configure Basic Settings

Fill in the registration form:

1. **Name**: `Autopilot Hash Import Tool`
   - This name will be visible to users during consent
   - You can use any descriptive name you prefer

2. **Supported account types**: Select one of the following:
   - ✅ **Accounts in this organizational directory only** (Single tenant) - **RECOMMENDED**
   - This ensures only users in your organization can use the tool

3. **Redirect URI**:
   - Platform: Select **Public client/native (mobile & desktop)**
   - URI: Enter `http://localhost`
   - This is required for interactive authentication

4. Click **Register**

### Step 4: Copy Important IDs

After registration, you'll see the **Overview** page:

1. **Copy the Application (client) ID**
   - This is a GUID (e.g., `12345678-1234-1234-1234-123456789abc`)
   - Save this to a text file - you'll need it later

2. **Copy the Directory (tenant) ID**
   - Also a GUID format
   - Save this as well

💡 **Tip**: Keep these IDs secure but accessible - you'll paste them into the AppConfig.json file.

---

## Part 2: Configure API Permissions

### Step 5: Add Microsoft Graph Permissions

1. In your app registration, click **API permissions** in the left menu
2. Click **+ Add a permission**
3. Select **Microsoft Graph**
4. Select **Delegated permissions**

### Step 6: Add Required Permissions

Search for and add these permissions:

#### Permission 1: DeviceManagementServiceConfig.ReadWrite.All
1. In the search box, type: `DeviceManagement`
2. Expand **DeviceManagementServiceConfig**
3. Check ✅ **DeviceManagementServiceConfig.ReadWrite.All**

#### Permission 2: User.Read
1. In the search box, type: `User`
2. Expand **User**
3. Check ✅ **User.Read**

4. Click **Add permissions**

### Step 7: Grant Admin Consent

⚠️ **IMPORTANT**: Admin consent is required for these permissions.

1. Click **✓ Grant admin consent for [Your Organization Name]**
2. Click **Yes** in the confirmation dialog
3. Wait for the status to change - all permissions should show:
   - **Status**: Green checkmark with "Granted for [Your Organization]"

💡 **Why is admin consent needed?**
- DeviceManagementServiceConfig.ReadWrite.All requires admin consent
- This is a one-time action that allows all users to use the app

---

## Part 3: Enable Public Client Flow

### Step 8: Configure Authentication Settings

1. Click **Authentication** in the left menu
2. Scroll down to **Advanced settings** section
3. Under **Allow public client flows**:
   - Set **Enable the following mobile and desktop flows** to **Yes**
4. Click **Save** at the top

💡 **Why enable public client flows?**
- This allows the PowerShell application to perform interactive authentication
- Required for delegated authentication without client secrets

---

## Part 4: Configure the Tool

### Step 9: Update AppConfig.json

1. Navigate to your tool folder: `AutopilotHashImportTool`
2. Open `Config\AppConfig.json` in a text editor (Notepad, VS Code, etc.)
3. Replace the placeholders with your IDs:

```json
{
    "ClientId": "PASTE_YOUR_APPLICATION_CLIENT_ID_HERE",
    "TenantId": "PASTE_YOUR_DIRECTORY_TENANT_ID_HERE",
    "GraphApiVersion": "beta"
}
```

**Example** (with dummy values):
```json
{
    "ClientId": "12345678-1234-1234-1234-123456789abc",
    "TenantId": "87654321-4321-4321-4321-cba987654321",
    "GraphApiVersion": "beta"
}
```

4. Save the file

---

## Part 5: Verify User Permissions

### Step 10: Assign Appropriate Roles

Users who will use this tool need **Intune permissions**. Assign one of these roles:

#### Option 1: Built-in Roles (Recommended)

Assign via **Entra ID** > **Roles and administrators**:

1. **Windows Autopilot Deployment Administrator** (Least Privilege) ✅
   - Can import and manage Autopilot devices
   - Cannot modify other Intune settings

2. **Intune Administrator** (Full Intune Access)
   - Full access to all Intune features
   - Use if user needs broader permissions

3. **Cloud Device Administrator**
   - Can manage devices and Autopilot
   - Broader than Autopilot admin but less than Intune admin

#### Option 2: Custom RBAC Role (Advanced)

If you need granular control, create a custom role with only:
- `Microsoft.Intune/ManagedDevices/Read`
- `Microsoft.Intune/ImportedWindowsAutopilotDeviceIdentities/Create`
- `Microsoft.Intune/ImportedWindowsAutopilotDeviceIdentities/Read`

### Step 11: Assign Users to Roles

1. Go to **Entra ID** > **Roles and administrators**
2. Search for **Windows Autopilot Deployment Administrator**
3. Click on the role
4. Click **+ Add assignments**
5. Search for and select the users who need access
6. Click **Add**

---

## Part 6: Test the Configuration

### Step 12: Run the Tool

1. Open PowerShell
2. Navigate to the tool folder:
   ```powershell
   cd "D:\Proj\AutopilotHashImportTool"
   ```
3. Run the tool:
   ```powershell
   .\AutopilotHashImportTool.ps1
   ```

### Step 13: Test Authentication

1. Click **🔐 Sign In to Microsoft Graph**
2. A browser window will open (or a pop-up will appear)
3. Enter your credentials
4. Accept the consent prompt (first time only)
5. You should see: "Status: ✓ Signed in as [your-email]"

✅ **Success**: If you see your email, authentication is working!

❌ **If authentication fails**, check:
- ClientId and TenantId are correct in AppConfig.json
- Admin consent was granted
- User has appropriate Intune role
- Internet connectivity is working

---

## Part 7: Optional - Advanced Configuration

### Restrict App Access (Optional)

By default, all users in your tenant can authenticate. To restrict access:

1. Go to **Entra ID** > **Enterprise applications**
2. Search for "Autopilot Hash Import Tool"
3. Click on the application
4. Go to **Properties**
5. Set **Assignment required?** to **Yes**
6. Click **Save**
7. Go to **Users and groups**
8. Click **+ Add user/group**
9. Select only the users/groups who should have access
10. Click **Assign**

### Enable Conditional Access (Optional)

Require MFA or compliant devices:

1. Go to **Entra ID** > **Security** > **Conditional Access**
2. Click **+ New policy**
3. Configure:
   - **Name**: Autopilot Tool - Require MFA
   - **Users**: Select users who use the tool
   - **Cloud apps**: Select "Autopilot Hash Import Tool"
   - **Grant**: Require multi-factor authentication
4. Enable policy and click **Create**

### Monitor Sign-ins

Track who's using the tool:

1. Go to **Entra ID** > **Monitoring** > **Sign-in logs**
2. Filter by **Application**: "Autopilot Hash Import Tool"
3. Review sign-in attempts, success/failure, locations, etc.

---

## Troubleshooting

### Issue: "AADSTS650052: The app needs access to a service"

**Solution**: Admin consent was not granted
- Go back to **API permissions**
- Click **Grant admin consent for [Your Organization]**

### Issue: "AADSTS7000218: The request body must contain the following parameter: 'client_assertion'"

**Solution**: Public client flow not enabled
- Go to **Authentication**
- Set **Allow public client flows** to **Yes**

### Issue: "Insufficient privileges to complete the operation"

**Solution**: User doesn't have Intune permissions
- Assign **Windows Autopilot Deployment Administrator** role
- OR assign **Intune Administrator** role

### Issue: "Invalid client"

**Solution**: ClientId is incorrect
- Verify the ClientId in AppConfig.json matches the Application (client) ID in Azure

### Issue: "AADSTS50020: User account from identity provider does not exist in tenant"

**Solution**: Wrong tenant ID
- Verify the TenantId in AppConfig.json matches your Directory (tenant) ID

---

## Security Best Practices

✅ **DO:**
- Grant least privilege roles (Windows Autopilot Deployment Administrator)
- Enable Conditional Access policies
- Monitor sign-in logs regularly
- Review API permissions periodically
- Keep the tool files in a secure location

❌ **DON'T:**
- Share ClientId/TenantId publicly (they're not secrets but should be controlled)
- Grant Global Administrator role just for this tool
- Allow unauthenticated access
- Disable MFA for tool users

---

## Quick Reference

### Required Azure AD Roles (for setup)
- Global Administrator OR Application Administrator

### Required Intune Roles (for users)
- Windows Autopilot Deployment Administrator (minimum)
- Intune Administrator (alternative)
- Cloud Device Administrator (alternative)

### Required API Permissions
- `DeviceManagementServiceConfig.ReadWrite.All` (Delegated)
- `User.Read` (Delegated)

### Configuration Files
- `Config\AppConfig.json` - Contains ClientId and TenantId

---

## Support

If you encounter issues:

1. **Check the status log** in the tool for detailed errors
2. **Review sign-in logs** in Entra ID
3. **Verify permissions** in the app registration
4. **Test with a different user** to rule out user-specific issues
5. **Check Intune service health** in Microsoft 365 admin center

---

## Summary Checklist

Use this checklist to verify your setup:

- [ ] App registered in Entra ID
- [ ] Application (client) ID copied
- [ ] Directory (tenant) ID copied
- [ ] Microsoft Graph API permissions added:
  - [ ] DeviceManagementServiceConfig.ReadWrite.All
  - [ ] User.Read
- [ ] Admin consent granted (green checkmarks visible)
- [ ] Public client flows enabled in Authentication
- [ ] AppConfig.json updated with ClientId and TenantId
- [ ] Users assigned appropriate Intune roles
- [ ] Tool tested successfully with authentication
- [ ] Sign-in logs reviewed

---

**Last Updated**: October 21, 2025  
**Version**: 1.0.0

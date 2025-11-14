# Azure Key Vault Setup Guide

This document provides step-by-step instructions for configuring Azure Key Vault to securely store the DigitalOwl API key.

## Prerequisites

- Azure subscription
- Azure CLI installed (optional but recommended)
- Access to create Azure resources
- .NET Framework 4.8 runtime

## Step 1: Create Azure Key Vault

### Using Azure Portal

1. Log in to [Azure Portal](https://portal.azure.com)
2. Click **Create a resource** > **Security** > **Key Vault**
3. Fill in the required information:
   - **Resource Group**: Create new or select existing
   - **Key Vault Name**: Choose a unique name (e.g., `digitalowl-keyvault`)
   - **Region**: Select your preferred region
   - **Pricing Tier**: Standard (recommended)
4. Click **Review + Create** > **Create**

### Using Azure CLI

```bash
# Login to Azure
az login

# Create resource group (if needed)
az group create --name DigitalOwlRG --location eastus

# Create Key Vault
az keyvault create \
  --name digitalowl-keyvault \
  --resource-group DigitalOwlRG \
  --location eastus
```

## Step 2: Store the API Key in Key Vault

### Using Azure Portal

1. Navigate to your Key Vault
2. Go to **Secrets** in the left menu
3. Click **+ Generate/Import**
4. Fill in:
   - **Upload options**: Manual
   - **Name**: `DigitalOwlApiKey`
   - **Value**: Your actual API key
5. Click **Create**

### Using Azure CLI

```bash
# Store the secret
az keyvault secret set \
  --vault-name digitalowl-keyvault \
  --name DigitalOwlApiKey \
  --value "your-actual-api-key-here"
```

## Step 3: Configure Access Permissions

The application uses `DefaultAzureCredential` which supports multiple authentication methods. Choose the appropriate method for your environment:

### Option A: Managed Identity (Recommended for Azure-hosted applications)

If running on Azure VM, App Service, or Azure Functions:

1. Enable Managed Identity on your Azure resource
2. Grant Key Vault access:

```bash
# Get the managed identity principal ID
PRINCIPAL_ID=$(az vm identity show --name YourVMName --resource-group YourRG --query principalId -o tsv)

# Grant access to Key Vault
az keyvault set-policy \
  --name digitalowl-keyvault \
  --object-id $PRINCIPAL_ID \
  --secret-permissions get list
```

### Option B: Service Principal (For on-premises or CI/CD)

1. Create a Service Principal:

```bash
az ad sp create-for-rbac --name "DigitalOwlApp" --skip-assignment
```

Save the output values (appId, password, tenant).

2. Grant Key Vault access:

```bash
az keyvault set-policy \
  --name digitalowl-keyvault \
  --spn <appId from previous step> \
  --secret-permissions get list
```

3. Set environment variables on the machine where the application runs:

```powershell
# Windows PowerShell
[System.Environment]::SetEnvironmentVariable('AZURE_CLIENT_ID', '<appId>', 'Machine')
[System.Environment]::SetEnvironmentVariable('AZURE_TENANT_ID', '<tenant>', 'Machine')
[System.Environment]::SetEnvironmentVariable('AZURE_CLIENT_SECRET', '<password>', 'Machine')
```

```bash
# Linux/Mac
export AZURE_CLIENT_ID="<appId>"
export AZURE_TENANT_ID="<tenant>"
export AZURE_CLIENT_SECRET="<password>"
```

### Option C: Azure CLI Authentication (For development)

1. Install Azure CLI
2. Run `az login`
3. Grant yourself access to the Key Vault:

```bash
# Get your user principal ID
USER_ID=$(az ad signed-in-user show --query id -o tsv)

# Grant access
az keyvault set-policy \
  --name digitalowl-keyvault \
  --object-id $USER_ID \
  --secret-permissions get list
```

### Option D: RBAC-based Access (Modern approach)

Instead of access policies, you can use RBAC:

```bash
# Enable RBAC on Key Vault
az keyvault update \
  --name digitalowl-keyvault \
  --enable-rbac-authorization true

# Assign role to user/service principal/managed identity
az role assignment create \
  --role "Key Vault Secrets User" \
  --assignee <user-email-or-object-id> \
  --scope /subscriptions/<subscription-id>/resourceGroups/<resource-group>/providers/Microsoft.KeyVault/vaults/digitalowl-keyvault
```

## Step 4: Update Application Configuration

Update `App.config` with your Key Vault details:

```xml
<appSettings>
    <!-- Azure Key Vault Configuration -->
    <add key="keyVaultUrl" value="https://digitalowl-keyvault.vault.azure.net/"/>
    <add key="keyVaultSecretName" value="DigitalOwlApiKey"/>
</appSettings>
```

Replace `digitalowl-keyvault` with your actual Key Vault name.

## Step 5: Install NuGet Packages

The following packages are required (already added to packages.config):

```bash
Install-Package Azure.Security.KeyVault.Secrets -Version 4.5.0
Install-Package Azure.Identity -Version 1.10.4
Install-Package Azure.Core -Version 1.36.0
Install-Package System.Memory -Version 4.5.5
Install-Package System.Text.Json -Version 8.0.0
Install-Package System.Threading.Tasks.Extensions -Version 4.5.4
Install-Package Microsoft.Bcl.AsyncInterfaces -Version 8.0.0
```

Or restore packages:

```bash
nuget restore
```

## Authentication Flow

The application uses `DefaultAzureCredential` which attempts authentication in this order:

1. **Environment Variables** - AZURE_CLIENT_ID, AZURE_TENANT_ID, AZURE_CLIENT_SECRET
2. **Managed Identity** - When running on Azure resources
3. **Visual Studio** - Uses logged-in Visual Studio account
4. **Azure CLI** - Uses `az login` credentials
5. **Azure PowerShell** - Uses PowerShell login

## Troubleshooting

### Error: "Authentication/Authorization failed" (401/403)

**Solution:**
- Verify the application has proper access to Key Vault
- Check Access Policies or RBAC roles are correctly configured
- Ensure you're authenticated (run `az login` for development)

### Error: "Secret not found" (404)

**Solution:**
- Verify the secret name matches exactly: `DigitalOwlApiKey`
- Check the Key Vault URL is correct
- Ensure the secret exists: `az keyvault secret show --vault-name digitalowl-keyvault --name DigitalOwlApiKey`

### Error: "DefaultAzureCredential failed to retrieve token"

**Solution:**
- Ensure at least one authentication method is configured
- For development: Run `az login`
- For production: Set up Managed Identity or Service Principal with environment variables
- Check the logs for specific authentication errors

### Verify Access

Test access to your Key Vault:

```bash
# Test reading the secret
az keyvault secret show \
  --vault-name digitalowl-keyvault \
  --name DigitalOwlApiKey \
  --query value -o tsv
```

## Security Best Practices

1. **Never commit secrets to source control** - The API key is now in Key Vault, not in code
2. **Use Managed Identity** - Preferred method for Azure-hosted applications (no credentials needed)
3. **Rotate secrets regularly** - Update the Key Vault secret periodically
4. **Audit access** - Enable Key Vault logging to monitor secret access
5. **Limit permissions** - Grant only "Get" and "List" permissions, not "Set" or "Delete"
6. **Use separate Key Vaults** - Different environments (dev, staging, prod) should use different Key Vaults

## Monitoring and Logging

Enable diagnostic logging for Key Vault:

```bash
az monitor diagnostic-settings create \
  --resource /subscriptions/<subscription-id>/resourceGroups/<resource-group>/providers/Microsoft.KeyVault/vaults/digitalowl-keyvault \
  --name KeyVaultLogs \
  --logs '[{"category": "AuditEvent", "enabled": true}]' \
  --workspace /subscriptions/<subscription-id>/resourceGroups/<resource-group>/providers/Microsoft.OperationalInsights/workspaces/<workspace-name>
```

## Migration from Word Document

If you're migrating from the old Word document approach:

1. **Backup** - Keep a copy of your `key.docx` file temporarily
2. **Store in Key Vault** - Follow Step 2 to store the key
3. **Test** - Run the application and verify it retrieves the key successfully
4. **Remove** - Once verified, you can delete the `key.docx` file
5. **Update** - The `keyFile` setting in App.config is now deprecated

## Cost Considerations

Azure Key Vault pricing (as of 2024):

- **Standard Tier**:
  - $0.03 per 10,000 operations
  - Secret operations: Get, List, Set, Delete
  - This application performs minimal operations (typically 1-2 per run)
  - Expected monthly cost: < $0.50

## Additional Resources

- [Azure Key Vault Documentation](https://docs.microsoft.com/azure/key-vault/)
- [DefaultAzureCredential Overview](https://docs.microsoft.com/dotnet/api/azure.identity.defaultazurecredential)
- [Key Vault Best Practices](https://docs.microsoft.com/azure/key-vault/general/best-practices)

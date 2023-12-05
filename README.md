# HAE_msGraphBackendUtilities
Back-end utility functions and model classes for working with Sharepoint Lists through the Microsoft Graph API.

All functions require a client-credentials configuration for use. This will require an app registration in Azure Entre ID. Utilities authenticate using an app ID, client credentials password, and an OAuth 2.0 Token with some attached scope. Generally, apps developed using these functions are intended to be Background/Deamon processes, and will require Application-level API permissions as opposed to delegated.

A function is provided to abstract the client secret for azure functions using environment variables via azure keyvault.
    this is implemented via https://azure.microsoft.com/en-us/blog/simplifying-security-for-serverless-and-web-apps-with-azure-functions-and-app-service/
    tutorial at: https://markheath.net/post/managed-identity-key-vault-azure-functions

For connecting to sharepoint using the client credentials flow, we highly reccomend the Sites.Selected Application-Level scope, and to add in read, write roles for just an individual sharepoint site when working with lists in order to provide the highest level of security possible.

Here are example MS Graph and PnP Powershell cmdlets that can be used to add red, write sp roles for your app registration.

    1. HTTP POST to MS Graph (With proper permissions):
    https://devblogs.microsoft.com/microsoft365dev/controlling-app-access-on-specific-sharepoint-site-collections/
    
    ------------------------------------
  
    POST https://graph.microsoft.com/v1.0/sites/{site_id}/permissions
  
    Content-Type: application/json
  
    {
  
      "roles": ["read","write"],
  
      "grantedToIdentities": [{
  
        "application": {
  
          "id": "Application (Client) ID From App Registration",
  
          "displayName": "Foo App"
  
        }
  
      }]
  
    }
  
    ------------------------------------
  
    2. PowerShell Cmdlet:
    https://pnp.github.io/powershell/cmdlets/Grant-PnPAzureADAppSitePermission.html

# HAE_msGraphBackendUtilities
Back-end utility functions and model classes for working with Sharepoint Lists through the Microsoft Graph API.

All functions require a client-credentials configuration for use. This will require an app registration in Azure Entre ID. Utilities authenticate using an app ID, client credentials password, and an OAuth 2.0 Token with some attached scope. Generally, apps developed using these functions are intended to be Background/Deamon processes, and will require Application-level API permissions as opposed to delegated.

A function is provided to abstract the client secret for azure functions using environment variables via azure keyvault.
    this is implemented via https://azure.microsoft.com/en-us/blog/simplifying-security-for-serverless-and-web-apps-with-azure-functions-and-app-service/
    tutorial at: https://markheath.net/post/managed-identity-key-vault-azure-functions

For connecting to sharepoint using the client credentials flow, we highly reccomend the Sites.Selected Application-Level scope, and to add in read, write roles for just an individual sharepoint site when working with lists in order to provide the highest level of security possible.

Here are example MS Graph and PnP Powershell cmdlets that can be used to add read, write sp roles for your app registration.

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

Here is an example client usage of the utilities:

    // creating a new client object
    MS_Graph_Deamon_Client msGraph = new MS_Graph_Deamon_Client(clientId, "msgraph_app_creds", tenantGuid);

    // Example GET:
    // this creates url: https://graph.microsoft.com/v1.0/sites/{siteid}/lists/{listid}/items?$expand=fields($select=id,EMail)&?$select=id
    // '50' sets max pages to grab to 50
    var getItemsResult = await GetSpListItems_MsGraph_V1(sp_siteid, sp_listid, oAuthToken, new List<string>(){ "id", "EMail", }, new List<string>(){ "id" }, 50);
    string getItemsJSON = (string)getItemsResult.Value;
    // deserialize JSON into your BYO model class
    List<UpdateItemMetaData> updateItems = JsonConvert.DeserializeObject<List<UpdateItemMetaData>>(getItemsJSON);

    // Example GET for user lookup ids in a site:
    var userListIdsResult = await GetSpListItems_MsGraph_V1(sp_siteid, sp_user, oAuthToken, new List<string>() { "id", "EMail", }, new List<string>() { "id" }, 10);
    string userListIdsJSON = (string)userListIdsResult.Value;
    // deserialize JSON into provided UserLookupMetaData model class
    List<UserLookupMetaData> userIdItems = JsonConvert.DeserializeObject<List<UserLookupMetaData>>(userListIdsJSON);

    // get a map of lookup ids
    Dictionary<string,string> userIdMap  = new Dictionary<string, string>();
    foreach (UserLookupMetaData userLookup in userIdItems)
    {
        foreach (UserLookupItem userIdItem in userLookup.lookup_id_result)
        {
            foreach (User user in data.users)
            {
                if (user.userEMail == userIdItem.lookup_data.email)
                {
                    userIdMap.Add(user.userEMail, userIdItem.id);
                }
            }
        }
    }

    // Example PATCH:
    var patchItemsResult = await PatchToSpList_MsGraph_V1(new Dictionary<string, string>(){ {"311", "{updateJson}"}, {"312", "{updateJson}"} }, sp_siteid, sp_listid, oAuthToken);
    if (patchItemsResult.DeclaredType != typeof(string))
    {
        // If the assigned work result is not of the expected type, return a BadRequest response with the result.
        return new BadRequestObjectResult($"PATCH Request: {patchItemsResult.Value}");
    }

    // Example POST:
    var postToSPResult = await PostToSpList_MsGraph_V1(new List<SharePointNewListItemModel>(){ new SharePointNewListItemModel() }, sp_siteid, sp_listid, oAuthToken);
    if (postToSPResult.DeclaredType != typeof(string))
    {
        // If the assigned work result is not of the expected type, return a BadRequest response with the result.
        return new BadRequestObjectResult($"POST Request: {postToSPResult.Value}");
    }
    

/// <summary>
/// Represents metadata for the result of a user lookup operation.
/// </summary>
[Serializable]
public class UserLookupMetaData
{
    [JsonProperty("@odata.context")] public string odata_context { get; set; }
    [JsonProperty("value")] public IList<UserLookupItem> lookup_id_result { get; set; }
    public UserLookupMetaData() { }
}

/// <summary>
/// Represents an item in the user lookup result.
/// </summary>
[Serializable]
public class UserLookupItem
{
    [JsonProperty("@odata.etag")] public string etag { get; set; }
    [JsonProperty("id")] public string id { get; set; }
    [JsonProperty("fields@odata.context")] public string fields_odata_context { get; set; }
    [JsonProperty("fields")] public UserLookupData lookup_data { get; set; }
    public UserLookupItem() { }
}

/// <summary>
/// Represents the data associated with a user lookup item.
/// </summary>
[Serializable]
public class UserLookupData
{
    [JsonProperty("@odata.etag")] public string etag { get; set; }
    [JsonProperty("id")] public string id { get; set; }
    [JsonProperty("EMail")] public string email { get; set; }
    public UserLookupData() { }
}

/// <summary>
/// Represents the response from an OAuth token request.
/// </summary>
[Serializable]
public class oAuthResponse
{
    [JsonProperty("token_type")] public string token_type { get; set; }
    [JsonProperty("expires_in")] public int expires_in { get; set; }
    [JsonProperty("ext_expires_in")] public int ext_expires_in { get; set; }
    [JsonProperty("access_token")] public string access_token { get; set; }

    public oAuthResponse() { }
}

/// <summary>
/// Provides custom HTTP Client methods for interacting with Microsoft Graph API using client credentials OAuth 2.0 flow
/// </summary>
public class MS_Graph_Deamon_Client{

    public MS_Graph_Deamon_Client() { };

    /// <summary>
    /// Sets the client secret variable to the app registration client secret stored in an Azure Key Vault.
    /// </summary>
    /// <param name="az_app_clientsecret">Reference to the Azure app client secret.</param>
    /// <param name="secretName">Name of the secret in Azure Key Vault.</param>
    static void Get_az_KV_Secret(ref string az_app_clientsecret, string secretName)
    {
        // -----------------------------------------------------------------------------------------------------------------------------------------------------
        // Method - Get_az_KV_Secret
        // -----------------------------------------------------------------------------------------------------------------------------------------------------
        /*

            implemented via https://azure.microsoft.com/en-us/blog/simplifying-security-for-serverless-and-web-apps-with-azure-functions-and-app-service/
            tutorial at: https://markheath.net/post/managed-identity-key-vault-azure-functions

            Sets the client secret variable to the app registration client secret stored in an azure key vault,
            and accessed via the Environment Variables.

        */

        az_app_clientsecret = Uri.EscapeDataString(Environment.GetEnvironmentVariable(secretName, EnvironmentVariableTarget.Process));
    }

    public async Task<string> GetResponseAsync(HttpClient client, HttpRequestMessage request)
    {
        HttpResponseMessage response = await client.SendAsync(request);
        return await response.Content.ReadAsStringAsync();
    }

    /// <summary>
    /// Asynchronously retrieves an OAuth access token using client credentials.
    /// </summary>
    /// <param name="clientid">The client ID of the application requesting the token.</param>
    /// <param name="clientsecret">The client secret associated with the application.</param>
    /// <param name="tenantguid">The Azure AD tenant's GUID.</param>
    /// <param name="scope">The requested scope for the access token.</param>
    /// <returns>An ObjectResult containing the OAuth access token if successful.</returns>
    public async Task<ObjectResult> GetOauthAccessToken(string clientid, string clientsecret, string tenantguid, string scope)
    {
        // -----------------------------------------------------------------------------------------------------------------------------------------------------
        // Method - GetOauthAccessToken  
        // -----------------------------------------------------------------------------------------------------------------------------------------------------
        /*

            Retrieves an OAuth access token using client credentials.

            Parameters:
                clientid (string): The client ID of the application requesting the token.
                clientsecret (string): The client secret associated with the application.
                tenantguid (string): The Azure AD tenant's GUID.
                scope (string): The requested scope for the access token.

            Returns:
                An ObjectResult containing the OAuth access token if successful.
                - Returns an AcceptedResult with the access token on success (HTTP 200 OK).
                - Returns a BadRequestObjectResult on failure with an error message.

            Remarks:
                This method asynchronously sends a request to the Microsoft Identity platform
                to obtain an OAuth access token using client credentials (client ID and client secret).

        */

        try
        {
            // Create an HTTP client for making the OAuth token request.
            HttpClient authClient = new HttpClient();
            authClient.DefaultRequestHeaders.Host = "login.microsoftonline.com:443";
            authClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            // Construct the OAuth token request.
            HttpRequestMessage request = new HttpRequestMessage()
            {
                RequestUri = new Uri($"https://login.microsoftonline.com/{tenantguid}/oauth2/v2.0/token"),
                Method = HttpMethod.Post,
                Content = new FormUrlEncodedContent(
                    new Dictionary<string, string>()
                    {
                        { "client_id", clientid },
                        { "scope", scope },
                        { "client_secret", clientsecret },
                        { "grant_type", "client_credentials" }
                    }
                ),
            };

            // Send the OAuth token request and await the response.
            HttpResponseMessage response = await authClient.SendAsync(request);
            var responseMessage = await response.Content.ReadAsStringAsync();
            HttpStatusCode status = response.StatusCode;

            // Throw an exception with details if the request fails.
            if (status != HttpStatusCode.OK) { throw new Exception($"{status} : {responseMessage}"); }

            // Parse the successful response and return the access token.
            oAuthResponse responseObject = JsonConvert.DeserializeObject<oAuthResponse>(responseMessage);
            return new AcceptedResult(location: null, value: responseObject.access_token);
        }
        catch (Exception ex)
        {
            // Return a BadRequestObjectResult with the exception message on failure.
            return new BadRequestObjectResult(ex.Message);
        }
    }

    /// <summary>
    /// Asynchronously retrieves items from a SharePoint list using Microsoft Graph API version 1.0.
    /// </summary>
    /// <param name="siteid">The ID of the SharePoint site where the list is located.</param>
    /// <param name="listid">The ID of the SharePoint list to retrieve items from.</param>
    /// <param name="oAuthToken">The OAuth access token for authentication.</param>
    /// <param name="itemFields">Optional: List of specific fields to retrieve for each item.</param>
    /// <param name="responseSelectFields">Optional: List of fields to include in the response.</param>
    /// <param name="paginationSearchLimit">Optional: Limit the number of items retrieved per request.</param>
    /// <returns>An ObjectResult containing the combined response from the Microsoft Graph API for retrieved items.</returns>
    public async Task<ObjectResult> GetSpListItems_MsGraph_V1(string siteid, string listid, string oAuthToken, List<string> itemFields = null, List<string> responseSelectFields = null, int paginationSearchLimit = 50)
    {
        // -----------------------------------------------------------------------------------------------------------------------------------------------------
        // Method - GetSpListItems_MsGraph_V1  
        // -----------------------------------------------------------------------------------------------------------------------------------------------------
        /*

            Implementation Details:
            - Constructs the Microsoft Graph API endpoint URI for retrieving items from a SharePoint list.
            - Sends HTTP GET requests to the Microsoft Graph API and paginates through the results.
            - Parses and combines the response for all retrieved items.
            - Handles errors and exceptions, returning appropriate results.

            Parameters:
            - siteid: The ID of the SharePoint site where the list is located.
            - listid: The ID of the SharePoint list to retrieve items from.
            - oAuthToken: The OAuth access token for authentication.
            - itemFields: Optional - List of specific fields to retrieve for each item.
            - responseSelectFields: Optional - List of fields to include in the response.
            - paginationSearchLimit: Optional - Limit the number of items retrieved per request.

            Returns:
            - An ObjectResult containing the combined response from the Microsoft Graph API for retrieved items.
            - AcceptedResult: Combined response on success (HTTP 200 OK).
            - BadRequestObjectResult: Error message on failure.

        */

        // Create an HTTP client for making the MS Graph SharePoint Item POST request.
        HttpClient getClient = new HttpClient();
        getClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        getClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", oAuthToken);

        string responseCombined = "";
        string fieldsExpandQuery = "";
        string fieldsSelectQuery = "";
        string fullUriString = $"https://graph.microsoft.com/v1.0/sites/{siteid}/lists/{listid}/items";

        if (itemFields != null)
        {
            fieldsExpandQuery += $"?expand=fields(select={string.Join(',', itemFields)})";
            fullUriString += fieldsExpandQuery;
        }

        if (responseSelectFields != null)
        {
            fieldsSelectQuery += $"?select={string.Join(',', responseSelectFields)}";
            if (fullUriString.Contains(fieldsExpandQuery)) { fullUriString += "&"; }
            fullUriString += fieldsSelectQuery;
        }

        try
        {
            for (int i = 0; i < paginationSearchLimit; i++)
            {
                // Construct the HTTP GET request.
                HttpRequestMessage request = new HttpRequestMessage()
                {
                    RequestUri = new Uri(fullUriString),
                    Method = HttpMethod.Get,
                };

                // Send the GET request to sharepoint over MS Graph and await the response.
                HttpResponseMessage response = await getClient.SendAsync(request);
                var responseMessage = await response.Content.ReadAsStringAsync();
                var status = response.StatusCode;

                // Throw an exception with details if the HTTP request fails.
                if (status != HttpStatusCode.OK) { throw new Exception($"{status} : {responseMessage}"); }

                // Get value data from response
                Regex regexValues = new Regex("\\\"[Vv][Aa][Ll][Uu][Ee]\\\"\\:\\s{0,}\\[(.*)\\]");
                Match matchValues = regexValues.Match(responseMessage);

                // Throw an exception with details if the regex parse fails.
                if (!matchValues.Success) { throw new Exception($"Unable to parse JSON objects via RegEx in GET Request response from Sharepoint."); }

                if (responseCombined != "")
                {
                    responseCombined += $", {matchValues.Groups[1].Value} ";
                }
                else
                {
                    responseCombined += $"[ {matchValues.Groups[1].Value} ";
                }

                if (responseMessage.Contains("@odata.nextLink"))
                {
                    // Get Pagination Link from @odata.nextLink property
                    Regex regexNextLink = new Regex("\\\"[@][Oo][Dd][Aa][Tt][Aa]\\.[Nn][Ee][Xx][Tt][Ll][Ii][Nn][Kk]\\\"\\:\\s{0,1}\\\"((?:.(?!\")){1,}[A-Za-z0-9])");
                    Match matchNextLink = regexNextLink.Match(responseMessage);

                    // break if no next page
                    if (!matchNextLink.Success) { break; }

                    fullUriString = matchNextLink.Groups[1].Value;
                }
            }
            responseCombined += " ]";
            // Return an AcceptedResult with the combined response from the Microsoft Graph API.
            return new AcceptedResult(location: null, value: responseCombined);
        }
        catch (Exception ex)
        {
            // Return an AcceptedResult with the combined response from the Microsoft Graph API.
            return new BadRequestObjectResult(ex.Message);
        }
    }

    /// <summary>
    /// Asynchronously posts new items to a SharePoint list using Microsoft Graph API version 1.0.
    /// </summary>
    /// <param name="listItems">The list of SharePoint list items to be posted.</param>
    /// <param name="siteid">The ID of the SharePoint site where the list is located.</param>
    /// <param name="listid">The ID of the SharePoint list where the items will be posted.</param>
    /// <param name="oAuthToken">The OAuth access token for authentication.</param>
    /// <returns>An ObjectResult containing the combined response from the Microsoft Graph API for all posted items.</returns>
    public async Task<ObjectResult> PostToSpList_MsGraph_V1(List<SharpointNewListItemModel> listItems, string siteid, string listid, string oAuthToken)
    {
        // -----------------------------------------------------------------------------------------------------------------------------------------------------
        // Method - PostToSpList_MsGraph_V1
        // -----------------------------------------------------------------------------------------------------------------------------------------------------
        /*

            Posts new items to a SharePoint list using Microsoft Graph API version 1.0.

            Parameters:
                listItems (List<SharpointNewListItemModel>): The list of SharePoint list items to be posted.
                siteid (string): The ID of the SharePoint site where the list is located.
                listid (string): The ID of the SharePoint list where the items will be posted.
                oAuthToken (string): The OAuth access token for authentication.

            Returns:
                An ObjectResult containing the combined response from the Microsoft Graph API for all posted items.
                - Returns an AcceptedResult with the combined response on success (HTTP 200 OK).
                - Returns a BadRequestObjectResult on failure with an error message.

            Remarks:
                This method asynchronously sends HTTP POST requests to add new items to a SharePoint list
                using Microsoft Graph API version 1.0. It requires a valid OAuth access token for authentication.

        */

        // Create an HTTP client for making the MS Graph SharePoint Item POST request.
        HttpClient postClient = new HttpClient();
        postClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        postClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", oAuthToken);

        string responseCombined = "";
        try
        {
            // Check if the list of SharePoint list items is provided and not null.
            if (listItems == null)
            {
                throw new Exception($"SP Post Request: No Sharepoint List Items were provided, or the object was null.");
            }

            List<HttpRequestMessage> httpRequestMessages = new List<HttpRequestMessage>();
            // Iterate through each SharePoint list item and send a POST request to add it to the list.
            foreach (SharpointNewListItemModel lim in listItems)
            {
                // Construct the HTTP POST request.
                //      * Please note that this method intends for the developer to implement a 
                //        JsonRequestContent Property on their implementation of their SharpointNewListItemModel
                //        class, the value of which will be a properly formatted JSON string for SP to read per:
                //        https://learn.microsoft.com/en-us/graph/api/listitem-create?view=graph-rest-1.0&tabs=http
                HttpRequestMessage request = new HttpRequestMessage()
                {
                    RequestUri = new Uri($"https://graph.microsoft.com/v1.0/sites/{siteid}/lists/{listid}/items"),
                    Method = HttpMethod.Post,
                    Content = new StringContent(
                        lim.JsonRequestContent,
                        Encoding.UTF8,
                        "application/json"
                    ),
                };

                httpRequestMessages.Add(request);
            }

            List<Task<string>> tasks = new List<Task<string>>();
            foreach (HttpRequestMessage request in httpRequestMessages)
            {
                tasks.Add(GetResponseAsync(postClient, request));
            }

            try
            {
                string[] resultBulkRequests = await Task.WhenAll(tasks.ToArray());
                foreach (string s in resultBulkRequests)
                {
                    if (responseCombined != "")
                    {
                        responseCombined += $", {s} ";
                    }
                    else
                    {
                        responseCombined += $"{s} ";
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"{ex.Message}");
            }
            // Return an AcceptedResult with the combined response from the Microsoft Graph API.
            if (responseCombined.Contains("BadRequest") || responseCombined.Contains("InternalServerError") || responseCombined.Contains("Invalid request"))
            {
                return new BadRequestObjectResult(responseCombined) { DeclaredType = typeof(Exception) };
            }

            return new AcceptedResult(location: null, value: responseCombined) { DeclaredType = typeof(string) };
        }
        catch (Exception ex)
        {
            // Return an AcceptedResult with the combined response from the Microsoft Graph API.
            return new BadRequestObjectResult(ex.Message) { DeclaredType = typeof(Exception) };
        }
    }

    /// <summary>
    /// Asynchronously updates SharePoint list items using Microsoft Graph API version 1.0.
    /// </summary>
    /// <param name="listItemUpdateContent">Dictionary of item IDs and corresponding JSON update content.</param>
    /// <param name="siteid">The ID of the SharePoint site where the list is located.</param>
    /// <param name="listid">The ID of the SharePoint list where the items will be updated.</param>
    /// <param name="oAuthToken">The OAuth access token for authentication.</param>
    /// <returns>An ObjectResult containing the combined response from the Microsoft Graph API for all updated items.</returns>
    public async Task<ObjectResult> PatchToSpList_MsGraph_V1(Dictionary<string, string> listItemUpdateContent, string siteid, string listid, string oAuthToken)
    {
        // -----------------------------------------------------------------------------------------------------------------------------------------------------
        // Method - PatchToSpList_MsGraph_V1
        // -----------------------------------------------------------------------------------------------------------------------------------------------------
        /*
        
            Implementation Details:
            - Constructs the Microsoft Graph API endpoint URI for updating items in a SharePoint list.
            - Sends HTTP PATCH requests to the Microsoft Graph API for each item in the provided update dictionary.
            - Parses and combines the response for all updated items.
            - Handles errors and exceptions, returning appropriate results.

            Parameters:
            - listItemUpdateContent: Dictionary of item IDs and corresponding JSON update content.
            - siteid: The ID of the SharePoint site where the list is located.
            - listid: The ID of the SharePoint list where the items will be updated.
            - oAuthToken: The OAuth access token for authentication.

            Returns:
            - An ObjectResult containing the combined response from the Microsoft Graph API for all updated items.
            - AcceptedResult: Combined response on success (HTTP 200 OK).
            - BadRequestObjectResult: Error message on failure.

        */

        // Create an HTTP client for making the MS Graph SharePoint Item POST request.
        HttpClient patchClient = new HttpClient();
        patchClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        patchClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", oAuthToken);

        string responseCombined = "";
        try
        {
            // Check if the list of SharePoint list items is provided and not null.
            if (listItemUpdateContent.Count < 1 || listItemUpdateContent == null)
            {
                throw new Exception($"SP Patch Request: No update fields were provided, or the object was null.");
            }

            List<HttpRequestMessage> httpRequestMessages = new List<HttpRequestMessage>();
            // Iterate through each SharePoint list item and send a POST request to add it to the list.
            foreach (KeyValuePair<string, string> idKey_contentValue in listItemUpdateContent)
            {
                // Construct the HTTP POST request.
                //      * Please note that this method intends for the developer to implement a 
                //        JsonRequestContent Property on their implementation of their SharpointNewListItemModel
                //        class, the value of which will be a properly formatted JSON string for SP to read per:
                //        https://learn.microsoft.com/en-us/graph/api/listitem-create?view=graph-rest-1.0&tabs=http
                HttpRequestMessage request = new HttpRequestMessage()
                {
                    RequestUri = new Uri($"https://graph.microsoft.com/v1.0/sites/{siteid}/lists/{listid}/items/{idKey_contentValue.Key}/fields"),
                    Method = HttpMethod.Patch,
                    Content = new StringContent(
                        idKey_contentValue.Value,
                        Encoding.UTF8,
                        "application/json"
                    ),
                };

                httpRequestMessages.Add(request);
            }

            List<Task<string>> tasks = new List<Task<string>>();
            foreach (HttpRequestMessage request in httpRequestMessages)
            {
                tasks.Add(GetResponseAsync(patchClient, request));
            }

            try
            {
                string[] resultBulkRequests = await Task.WhenAll(tasks.ToArray());
                foreach (string s in resultBulkRequests)
                {
                    if (responseCombined != "")
                    {
                        responseCombined += $", {s} ";
                    }
                    else
                    {
                        responseCombined += $"{s} ";
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"{ex.Message}");
            }
            // Return an AcceptedResult with the combined response from the Microsoft Graph API.
            if (responseCombined.Contains("BadRequest")||responseCombined.Contains("InternalServerError")||responseCombined.Contains("Invalid request"))
            {
                return new BadRequestObjectResult(responseCombined) { DeclaredType = typeof(Exception) };
            }
            return new AcceptedResult(location: null, value: responseCombined) { DeclaredType = typeof(string) };
        }
        catch (Exception ex)
        {
            // Return an AcceptedResult with the combined response from the Microsoft Graph API.
            return new BadRequestObjectResult(ex.Message) { DeclaredType = typeof(Exception) };
        }
    }
}


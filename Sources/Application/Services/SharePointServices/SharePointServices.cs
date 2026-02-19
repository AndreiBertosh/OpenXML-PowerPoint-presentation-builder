using Application.Services.SharePointServices.Interfaces;
using Application.Services.SharePointServices.Models;
using Azure.Core;
using Azure.Identity;
using Infrastructure.Common;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Identity.Client;
using System.Net;
using System.Net.Http.Headers;
using System.Text.Json;

namespace Application.Services.SharePointServices
{
    /// <summary>
    /// Service for working with SharePoint files via Microsoft Graph API.
    /// </summary>
    public class SharePointServices : ISharePointServices
    {
        private readonly string _tenantId;
        private readonly string _clientId;
        private readonly string _clientSecret;
        private readonly string _siteHost;
        private readonly string _sitePath;
        private readonly string _folderPath;
        private readonly string _targetFolderName;
        private readonly string _webproxyAddress;
        private readonly bool _useWebproxy;
        private readonly ILogger<SharePointServices> _logger;

        private static readonly string[] GraphScopes = ["https://graph.microsoft.com/.default"];

        public SharePointServices(IOptions<AzureAppSettings> settings, ILogger<SharePointServices> logger)
        {
            var configuration = settings.Value;
            _tenantId = configuration.TenantId;
            _clientId = configuration.ClientId;
            _clientSecret = configuration.ClientSecret;
            _siteHost = configuration.SiteHost;
            _sitePath = configuration.SitePath;
            _folderPath = configuration.TemplatesFolderName;
            _targetFolderName = configuration.ResultFolderName;
            _webproxyAddress = configuration.WebProxyAddress;
            _useWebproxy =  configuration.UseWebProxy;
            _logger = logger;
        }

        #region Private Helpers

        private HttpClient CreateProxiedHttpClient()
        {
            var handler = new HttpClientHandler
            {
                //Proxy = new WebProxy("http://194.170.80.7:9480"),
                //UseProxy = true
                Proxy = new WebProxy(_webproxyAddress),
                UseProxy = _useWebproxy
            };

            return new HttpClient(handler, disposeHandler: true);
        }


        // Creates a reusable Graph client instance
        private GraphServiceClient CreateGraphClient()
        {
            var credential = new ClientSecretCredential(
                _tenantId,
                _clientId,
                _clientSecret);

            var httpClient = CreateProxiedHttpClient();

            return new GraphServiceClient(httpClient, credential, GraphScopes);
        }


        // Gets SiteId via Graph API with error handling
        private async Task<string> GetSiteIdAsync(GraphServiceClient client, CancellationToken cancellationToken)
        {
            try
            {
                Site? site = await client.Sites[$"{_siteHost}:/{_sitePath}"].GetAsync();
                return site?.Id ?? throw new SharePointServiceException("Site not found");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error while getting SiteId");
                throw new SharePointServiceException("Error while getting SiteId", ex);
            }
        }

        // Gets DriveId via Graph API with error handling
        private async Task<string> GetDriveIdAsync(GraphServiceClient client, string siteId, CancellationToken cancellationToken)
        {
            try
            {
                Drive? drive = await client.Sites[siteId].Drive.GetAsync(cancellationToken: cancellationToken);
                return drive?.Id ?? throw new SharePointServiceException("Drive not found");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error while getting DriveId");
                throw new SharePointServiceException("Error while getting DriveId", ex);
            }
        }

        // Finds a folder by name in the root of the drive
        private async Task<string?> FindFolderIdAsync(GraphServiceClient client, string driveId, string folderName, CancellationToken cancellationToken)
        {
            try
            {
                var rootItems = await client.Drives[driveId].Items["root"].Children.GetAsync(cancellationToken: cancellationToken);
                return rootItems?.Value?
                    .FirstOrDefault(item => item.Folder != null && item.Name.Equals(folderName, StringComparison.OrdinalIgnoreCase))?.Id;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error searching folder '{folderName}'");
                throw new SharePointServiceException($"Error searching folder '{folderName}'", ex);
            }
        }

        #endregion

        #region Business Methods

        /// <summary>
        /// Returns a dictionary of folder names and lists of files in each folder at the SharePoint root.
        /// </summary>
        public async Task<Dictionary<string, List<string>>> GetRootFilesAsync(CancellationToken cancellationToken)
        {
            var client = CreateGraphClient();
            string siteId = await GetSiteIdAsync(client, cancellationToken);
            string driveId = await GetDriveIdAsync(client, siteId, cancellationToken);

            var rootItems = await client.Drives[driveId].Items["root"].Children.GetAsync(cancellationToken: cancellationToken);
            Dictionary<string, List<string>> folderFiles = [];

            foreach (var item in rootItems?.Value ?? Enumerable.Empty<DriveItem>())
            {
                if (item.Folder != null && item.Id is not null)
                {
                    // Get the list of files in each folder
                    var files = await client.Drives[driveId].Items[item.Id].Children.GetAsync(cancellationToken: cancellationToken);
                    folderFiles[item.Name] = files?.Value?
                        .Where(f => f.File != null)
                        .Select(f => f.Name)
                        .ToList() ?? [];
                }
            }
            return folderFiles;
        }

        /// <summary>
        /// Downloads a file by name from a specific folder.
        /// </summary>
        public async Task<MemoryStream> DownloadFileByNameAsync(string folderName, string fileName, CancellationToken cancellationToken)
        {
            var client = CreateGraphClient();
            string siteId = await GetSiteIdAsync(client, cancellationToken);
            string driveId = await GetDriveIdAsync(client, siteId, cancellationToken);

            string filePath = string.IsNullOrEmpty(folderName) ? fileName : $"{folderName}/{fileName}";
            DriveItem? fileItem;
            try
            {
                fileItem = await client.Drives[driveId].Root.ItemWithPath(filePath).GetAsync(cancellationToken: cancellationToken);
                if (fileItem?.Id is null)
                {
                    throw new FileNotFoundException($"File '{fileName}' not found in folder '{folderName}'.");
                }
            }
            catch (ServiceException ex) when (ex.ResponseStatusCode == (int)System.Net.HttpStatusCode.NotFound)
            {
                _logger.LogWarning(ex, $"File not found: {fileName} in {folderName}");
                throw new FileNotFoundException($"File '{fileName}' in folder '{folderName}' not found.", ex);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Unexpected error in DownloadFileByNameAsync");
                throw;
            }

            try
            {
                // Download file content to a memory stream
                Stream? contentStream = await client.Drives[driveId].Items[fileItem.Id].Content.GetAsync();
                MemoryStream memoryStream = new();
                await contentStream.CopyToAsync(memoryStream, cancellationToken);
                memoryStream.Position = 0;
                return memoryStream;
            }
            catch (ServiceException ex)
            {
                _logger.LogError(ex, $"Failed to download file: {fileName}");
                throw new IOException($"Failed to download file '{fileName}'. Status: {ex.ResponseStatusCode}", ex);
            }
        }

        /// <summary>
        /// Uploads a file stream into a specified folder.
        /// </summary>
        public async Task<string> UploadFileAsync1(string folderPath, string fileName, Stream contentStream, CancellationToken cancellationToken)
        {
            var client = CreateGraphClient();
            string siteId = await GetSiteIdAsync(client, cancellationToken);
            string driveId = await GetDriveIdAsync(client, siteId, cancellationToken);
            string targetFolder = folderPath ?? _targetFolderName;
            string targetFolderId = await FindFolderIdAsync(client, driveId, targetFolder, cancellationToken)
                ?? throw new SharePointServiceException($"Folder '{targetFolder}' not found in SharePoint root.");

            // Create an upload session for large files
            UploadSession? uploadSession = await client.Drives[driveId]
                .Items[targetFolderId].ItemWithPath(fileName)
                .CreateUploadSession.PostAsync(new Microsoft.Graph.Drives.Item.Items.Item.CreateUploadSession.CreateUploadSessionPostRequestBody
                {
                    Item = new DriveItemUploadableProperties { Name = fileName }
                }, cancellationToken: cancellationToken);

            int maxChunkSize = 320 * 1024;
            LargeFileUploadTask<DriveItem> uploader = new(uploadSession, contentStream, maxChunkSize);

            try
            {
                var uploadResult = await uploader.UploadAsync(null, cancellationToken: cancellationToken);
                return uploadResult.UploadSucceeded
                    ? uploadResult.ItemResponse.WebUrl
                    : throw new SharePointServiceException("Upload session did not succeed.");
            }
            catch (ServiceException ex)
            {
                _logger.LogError(ex, "Error uploading file to SharePoint");
                throw new SharePointServiceException("Error uploading file to SharePoint.", ex);
            }
        }

        /// <summary>
        /// Returns dictionary with folder/subfolder names and lists of files in them for the given folder.
        /// </summary>
        public async Task<Dictionary<string, List<FileInfoDomainModel>>> GetFilesInFolderAsync(string folderName, CancellationToken cancellationToken)
        {
            var client = CreateGraphClient();
            string siteId = await GetSiteIdAsync(client, cancellationToken);
            string driveId = await GetDriveIdAsync(client, siteId, cancellationToken);
            string targetFolderId = await FindFolderIdAsync(client, driveId, folderName, cancellationToken)
                ?? throw new SharePointServiceException($"Folder '{folderName}' not found in SharePoint root.");

            var folderChildren = await client.Drives[driveId].Items[targetFolderId].Children.GetAsync(cancellationToken: cancellationToken);

            Dictionary<string, List<FileInfoDomainModel>> result = [];
            foreach (var item in folderChildren.Value)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (item.Folder != null)
                {
                    // Get files in the subfolder
                    List<FileInfoDomainModel> subfolderFiles = [] ;
                    var children = await client.Drives[driveId].Items[item.Id].Children.GetAsync(cancellationToken: cancellationToken);

                    foreach (var subItem in children.Value)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        if (subItem.File != null)
                        {
                            subfolderFiles.Add(new FileInfoDomainModel(subItem.Name, subItem.Size, subItem.LastModifiedDateTime));
                        }
                    }
                    result[item.Name] = subfolderFiles;
                }
                else if (item.File != null)
                {
                    // Add the file to the main folder
                    if (!result.ContainsKey(folderName))
                    {
                        result[folderName] = [];
                    }

                    result[folderName].Add(new FileInfoDomainModel(item.Name, item.Size, item.LastModifiedDateTime));
                }
            }
            return result;
        }

        /// <summary>
        /// Returns the direct download link for the specified file.
        /// </summary>
        public async Task<string> GetFileDownloadLinkAsync(string fileName, CancellationToken cancellationToken)
        {
            if (string.IsNullOrWhiteSpace(fileName))
            {
                throw new ArgumentException(null, nameof(fileName));
            }

            var client = CreateGraphClient();
            string siteId = await GetSiteIdAsync(client, cancellationToken);
            string driveId = await GetDriveIdAsync(client, siteId, cancellationToken);

            string targetFolderId = await FindFolderIdAsync(client, driveId, _targetFolderName, cancellationToken)
                ?? throw new SharePointServiceException("Target folder not found.");

            var folderChildren = await client.Drives[driveId].Items[targetFolderId].Children.GetAsync(cancellationToken: cancellationToken);
            var fileItem = folderChildren.Value.FirstOrDefault(i => i.File != null && i.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                ?? throw new SharePointServiceException($"File '{fileName}' not found.");

            // Get download link via AdditionalData
            return fileItem.AdditionalData != null && fileItem.AdditionalData.TryGetValue("@microsoft.graph.downloadUrl", out var urlObj)
                ? urlObj?.ToString() ?? throw new SharePointServiceException("No downloadUrl provided.")
                : throw new SharePointServiceException("Unable to retrieve download link for the file.");
        }

        /// <summary>
        /// Checks if the specified file exists in the given folder.
        /// </summary>
        public async Task<bool> DoesFileExistAsync(string folderName, string fileName, CancellationToken cancellationToken)
        {
            var client = CreateGraphClient();
            string siteId = await GetSiteIdAsync(client, cancellationToken);
            string driveId = await GetDriveIdAsync(client, siteId, cancellationToken);
            string folderToSearch = string.IsNullOrWhiteSpace(folderName) ? _targetFolderName : folderName;
            string? targetFolderId = await FindFolderIdAsync(client, driveId, folderToSearch, cancellationToken);

            if (targetFolderId is null)
            {
                return false;
            }

            // Check if file exists in the folder
            var folderChildren = await client.Drives[driveId].Items[targetFolderId].Children.GetAsync(cancellationToken: cancellationToken);
            return folderChildren?.Value?.Any(f => f.File != null && f.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) == true;
        }

        /// <summary>
        /// Deletes a file by name from a specified folder in SharePoint.
        /// </summary>

        public async Task DeleteFileAsync(string folderName, string fileName, CancellationToken cancellationToken)
        {
            var client = CreateGraphClient();
            string siteId = await GetSiteIdAsync(client, cancellationToken);
            string driveId = await GetDriveIdAsync(client, siteId, cancellationToken);

            // Reuse FindFolderIdAsync to get folder ID
            string targetFolderId = await FindFolderIdAsync(client, driveId, folderName, cancellationToken)
                ?? throw new SharePointServiceException($"Folder '{folderName}' not found in SharePoint root.");

            // Get folder children and find the file
            var folderChildren = await client.Drives[driveId].Items[targetFolderId].Children.GetAsync(cancellationToken: cancellationToken);
            var fileItem = folderChildren.Value
                .FirstOrDefault(item => item.File != null && item.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) ?? throw new SharePointServiceException($"File '{fileName}' not found in folder '{folderName}'.");
            
            try
            {
                // Delete the file by its ID
                await client.Drives[driveId].Items[fileItem.Id].DeleteAsync(cancellationToken: cancellationToken);
                _logger.LogInformation("Deleted file '{FileName}' from SharePoint folder '{FolderName}'.", fileName, folderName);
            }
            catch (ServiceException ex)
            {
                _logger.LogError(ex, "Error deleting file '{FileName}' from SharePoint folder '{FolderName}': {Message}", fileName, folderName, ex.Message);
                throw new SharePointServiceException($"Error deleting file '{fileName}' from folder '{folderName}'.", ex);
            }
        }

        /// <summary>
        /// Subscribes the provided Azure Function webhook URL to SharePoint change events ("created", "updated") for the _folderPath folder, for 180 days.
        /// </summary>
        public async Task<string> SubscribeWebHookForFolderAsync(string azureFunctionUrl, CancellationToken cancellationToken)
        {
            if (string.IsNullOrWhiteSpace(azureFunctionUrl))
                throw new ArgumentException("Webhook URL for Azure Function must not be empty.", nameof(azureFunctionUrl));

            var client = CreateGraphClient();
            string siteId = await GetSiteIdAsync(client, cancellationToken);
            string driveId = await GetDriveIdAsync(client, siteId, cancellationToken);

            // Find folderId (ItemId) for _folderPath
            string folderId = await FindFolderIdAsync(client, driveId, _folderPath, cancellationToken)
                ?? throw new SharePointServiceException($"Folder '{_folderPath}' not found for WebHook subscription.");

            // Calculate expiration date (max supported is 180 days for resource = drive/root/items/{id}, see Microsoft Docs)
            var expiration = DateTime.UtcNow.AddDays(180);

            var subscriptionPayload = new
            {
                changeType = "created,updated",
                notificationUrl = azureFunctionUrl,
                resource = $"/drives/{driveId}/items/{folderId}",
                expirationDateTime = expiration.ToString("o"),
                clientState = Guid.NewGuid().ToString()
            };

            // Get OAuth token for Graph API
            var graphApiUrl = "https://graph.microsoft.com/v1.0/subscriptions";
            var tokenRequestContext = new TokenRequestContext(GraphScopes);
            var credential = new ClientSecretCredential(_tenantId, _clientId, _clientSecret);
            var token = await credential.GetTokenAsync(tokenRequestContext, cancellationToken);

            using var httpClient = CreateProxiedHttpClient();
            httpClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token.Token);

            var jsonContent = new StringContent(
                JsonSerializer.Serialize(subscriptionPayload),
                System.Text.Encoding.UTF8,
                "application/json"
            );

            var response = await httpClient.PostAsync(graphApiUrl, jsonContent, cancellationToken);
            var responseContent = await response.Content.ReadAsStringAsync(cancellationToken);

            if (response.IsSuccessStatusCode)
            {
                _logger.LogInformation("Successfully created webhook subscription for folder '{FolderPath}' to URL '{WebhookUrl}'. Expiration: {Expiration}. Response: {Response}", _folderPath, azureFunctionUrl, expiration, responseContent);
                return responseContent;
            }
            else
            {
                _logger.LogError("Failed to create webhook subscription for folder '{FolderPath}': {Status} - {Response}", _folderPath, response.StatusCode, responseContent);
                throw new SharePointServiceException($"Failed to create webhook subscription. Status: {response.StatusCode}. Response: {responseContent}");
            }
        }

        #endregion

        #region testRegion

        public async Task<string> UploadFileAsync(string folderPath, string fileName, Stream contentStream, CancellationToken cancellationToken)
        {
            folderPath ??= _targetFolderName;

            // Step 1: Acquire token
            var app = ConfidentialClientApplicationBuilder.Create(_clientId)
                .WithClientSecret(_clientSecret)
                .WithAuthority(new Uri($"https://login.microsoftonline.com/{_tenantId}"))
                .Build();

            var result = await app.AcquireTokenForClient(["https://graph.microsoft.com/.default"]).ExecuteAsync(cancellationToken);
            string accessToken = result.AccessToken;

            // Step 2: Create HttpClient
            using var httpClient = CreateProxiedHttpClient();
            httpClient.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", accessToken);

            // Step 3: Resolve siteId
            string siteEndpoint = $"https://graph.microsoft.com/v1.0/sites/{_siteHost}:/{_sitePath}";
            var siteResponse = await httpClient.GetAsync(siteEndpoint, cancellationToken);
            siteResponse.EnsureSuccessStatusCode();
            string siteJson = await siteResponse.Content.ReadAsStringAsync(cancellationToken);
            using var siteDoc = System.Text.Json.JsonDocument.Parse(siteJson);
            string siteId = siteDoc.RootElement.GetProperty("id").GetString();

            // Step 4: Build upload endpoint
            // PUT /sites/{siteId}/drive/root:/{folderPath}/{fileName}:/content
            string uploadEndpoint = $"https://graph.microsoft.com/v1.0/sites/{siteId}/drive/root:/{folderPath}/{fileName}:/content";

            // Step 5: Send PUT request with stream
            using var streamContent = new StreamContent(contentStream);
            streamContent.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");

            var uploadResponse = await httpClient.PutAsync(uploadEndpoint, streamContent, cancellationToken);
            if (!uploadResponse.IsSuccessStatusCode)
            {
                throw new IOException(
                    $"Failed to upload file '{fileName}' to folder '{folderPath}'. Status: {uploadResponse.StatusCode}");
            }

            return uploadResponse.ReasonPhrase;
        }

        public async Task<MemoryStream> DownloadFileByNameAsync2(string folderName, string fileName, CancellationToken cancellationToken)
        {
            // Step 1: Acquire token
            var app = ConfidentialClientApplicationBuilder.Create(_clientId)
                .WithClientSecret(_clientSecret)
                .WithAuthority(new Uri($"https://login.microsoftonline.com/{_tenantId}"))
                .Build();

            var result = await app.AcquireTokenForClient(["https://graph.microsoft.com/.default"]).ExecuteAsync();
            string accessToken = result.AccessToken;

            // Step 2: Create HttpClient
            using var httpClient = CreateProxiedHttpClient();
            httpClient.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", accessToken);

            // Step 3: Resolve siteId
            string siteEndpoint = $"https://graph.microsoft.com/v1.0/sites/{_siteHost}:/{_sitePath}";
            var siteResponse = await httpClient.GetAsync(siteEndpoint);
            siteResponse.EnsureSuccessStatusCode();
            string siteJson = await siteResponse.Content.ReadAsStringAsync();
            using var siteDoc = JsonDocument.Parse(siteJson);
            string siteId = siteDoc!.RootElement.GetProperty("id").GetString();

            // Step 4: Get file metadata by path
            string fileEndpoint = $"https://graph.microsoft.com/v1.0/sites/{siteId}/drive/root:/{folderName}/{fileName}";
            var fileResponse = await httpClient.GetAsync(fileEndpoint);

            if (!fileResponse.IsSuccessStatusCode)
            {
                throw new FileNotFoundException(
                    $"File '{fileName}' in folder '{folderName}' not found or access denied. " +
                    $"Status: {fileResponse.StatusCode}");
            }

            string fileJson = await fileResponse.Content.ReadAsStringAsync();
            using var fileDoc = JsonDocument.Parse(fileJson);

            if (!fileDoc.RootElement.TryGetProperty("id", out var idProp))
            {
                throw new FileNotFoundException(
                    $"File '{fileName}' in folder '{folderName}' not found in Graph response.");
            }

            string fileId = idProp.GetString();

            // Step 5: Download file content
            string downloadEndpoint = $"https://graph.microsoft.com/v1.0/sites/{siteId}/drive/items/{fileId}/content";
            var downloadResponse = await httpClient.GetAsync(downloadEndpoint, HttpCompletionOption.ResponseHeadersRead);

            if (!downloadResponse.IsSuccessStatusCode)
            {
                throw new IOException(
                    $"Failed to download file '{fileName}'. Status: {downloadResponse.StatusCode}");
            }

            // Step 6: Copy into MemoryStream
            var memoryStream = new MemoryStream();
            using (var responseStream = await downloadResponse.Content.ReadAsStreamAsync())
            {
                await responseStream.CopyToAsync(memoryStream);
            }

            // Reset position to beginning
            memoryStream.Position = 0;

            return memoryStream;
        }

        #endregion
    }
}
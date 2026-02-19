using Application.Services.AzureServices.Interfaces;

using Azure;
using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;
using Azure.Storage.Sas;

using Infrastructure.Common;

using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace Application.Services.AzureServices
{
    public class AzureServices : IAzureServices
    {
        private readonly AzureBlobStorageSettings _settings;
        private readonly ILogger<AzureServices> _logger;
        private readonly string _connectionString;

        public AzureServices(IOptions<AzureBlobStorageSettings> settings, ILogger<AzureServices> logger)
        {
            _settings = settings.Value;
            _connectionString = settings.Value.ConnectionString;
            _logger = logger;
        }

        public async Task<bool> DeleteBlobAsync(string containerName, string blobName, CancellationToken cancellationToken)
        {
            try
            {
                // 1. Create a client for the Blob Storage service
                BlobServiceClient blobServiceClient = new(_connectionString);

                // 2. Get the client for the Blob Container
                BlobContainerClient containerClient = blobServiceClient.GetBlobContainerClient(containerName);

                // 3. Get the client for the specific Blob object
                BlobClient blobClient = containerClient.GetBlobClient(blobName);

                // 4. Attempt to delete the blob with cancellation token
                var response = await blobClient.DeleteIfExistsAsync(cancellationToken: cancellationToken);

                if (response.Value)
                {
                    _logger.LogInformation("Blob '{BlobName}' deleted successfully.", blobName);
                    return true;
                }
                else
                {
                    _logger.LogWarning("Blob '{BlobName}' does not exist or was already deleted.", blobName);
                    return false;
                }
            }
            catch (RequestFailedException ex)
            {
                _logger.LogError(ex, "Failed to delete blob '{BlobName}': {Message}", blobName, ex.Message);
                throw;
            }
        }

        public async Task<IEnumerable<string>> GetBlobsAsync(string containerName, CancellationToken cancellationToken)
        {
            try
            {
                // 1. Create a client for the Blob Storage service
                BlobServiceClient blobServiceClient = new(_connectionString);

                // 2. Get the client for the Blob Container
                BlobContainerClient containerClient = blobServiceClient.GetBlobContainerClient(containerName);
                await containerClient.CreateIfNotExistsAsync(cancellationToken: cancellationToken);

                // 3. Get Blobs
                List<string> blobNames = [];
                _logger.LogInformation("Getting List of Blobs from container '{ContainerName}'", containerName);

                await foreach (BlobItem blobItem in containerClient.GetBlobsAsync(cancellationToken: cancellationToken))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    blobNames.Add(blobItem.Name);
                }

                return blobNames;
            }
            catch (RequestFailedException ex) when (ex.Status >= 400 && ex.Status < 500)
            {
                // Log client-side errors (HTTP 4xx)
                _logger.LogWarning(ex, "Client error while trying to get List of Blob objects: {Message}", ex.Message);
                throw new Exception($"The list of Blob objects could not be retrieved. Error: {ex.Message}", ex);
            }
            catch (RequestFailedException ex)
            {
                // Log server-side errors or unexpected failures
                _logger.LogError(ex, "Error while trying to get List of Blob objects: {Message}", ex.Message);
                throw;
            }
        }

        public async Task<Stream> GetBlobStreamAsync(string container, string blobName, CancellationToken cancellationToken)
        {
            try
            {
                // 1. Create a client for the Blob Storage service
                BlobServiceClient blobServiceClient = new(_connectionString);

                // 2. Get the client for the Blob Container
                BlobContainerClient containerClient = blobServiceClient.GetBlobContainerClient(container);

                // Use cancellationToken for container creation
                await containerClient.CreateIfNotExistsAsync(cancellationToken: cancellationToken);

                // 3. Get the client for the specific Blob object
                BlobClient blobClient = containerClient.GetBlobClient(blobName);

                // 4. Open the Blob object as a stream
                _logger.LogInformation("Fetching Blob '{BlobName}' from container '{ContainerName}'", blobName, container);

                // Use cancellationToken for stream opening
                return await blobClient.OpenReadAsync(cancellationToken: cancellationToken);
            }
            catch (RequestFailedException ex) when (ex.Status >= 400 && ex.Status < 500)
            {
                // Log client-side errors (HTTP 4xx)
                _logger.LogWarning(ex, "Client error while trying to fetch Blob '{BlobName}': {Message}", blobName, ex.Message);
                throw new Exception($"Blob '{blobName}' could not be retrieved. Error: {ex.Message}", ex);
            }
            catch (RequestFailedException ex)
            {
                // Log server-side errors or unexpected failures
                _logger.LogError(ex, "Error while trying to fetch Blob '{BlobName}': {Message}", blobName, ex.Message);
                throw;
            }
        }

        public async Task<string> GetDownloadLinkAsync(string containerName, string blobName, CancellationToken cancellationToken)
        {
            var blobServiceClient = new BlobServiceClient(_connectionString);
            var containerClient = blobServiceClient.GetBlobContainerClient(containerName);

            // Pass cancellationToken to async call
            await containerClient.CreateIfNotExistsAsync(cancellationToken: cancellationToken);

            var blobClient = containerClient.GetBlobClient(blobName);

            // Pass cancellationToken to async call
            if (!await blobClient.ExistsAsync(cancellationToken: cancellationToken))
            {
                _logger.LogWarning("Blob '{BlobName}' does not exist. Cannot generate SAS link.", blobName);
                throw new FileNotFoundException($"Blob '{blobName}' not found.");
            }

            if (!blobClient.CanGenerateSasUri)
            {
                throw new InvalidOperationException("BlobClient cannot generate SAS URI. Make sure you're using a key-based credential.");
            }

            var sasBuilder = new BlobSasBuilder
            {
                BlobContainerName = containerName,
                BlobName = blobName,
                Resource = "b",
                ExpiresOn = DateTimeOffset.UtcNow.AddMinutes(60)
            };

            sasBuilder.SetPermissions(BlobSasPermissions.Read);

            var sasUri = blobClient.GenerateSasUri(sasBuilder);

            return sasUri.ToString();
        }
        public async Task<Stream> GetWritableBlobStreamAsync(string containerName, string blobName, CancellationToken cancellationToken)
        {
            try
            {
                // 1. Create a client for the Blob Storage service
                BlobServiceClient blobServiceClient = new (_connectionString);

                // 2. Get the client for the Blob Container
                BlobContainerClient containerClient = blobServiceClient.GetBlobContainerClient(containerName);
                await containerClient.CreateIfNotExistsAsync(cancellationToken: cancellationToken);

                // 3. Get the client for the specific Blob object
                BlobClient blobClient = containerClient.GetBlobClient(blobName);

                // 4. Create writable stream and download the Blob content into it
                _logger.LogInformation("Attempting to download Blob '{BlobName}' to a writable stream.", blobName);

                MemoryStream writableStream = new();
                await blobClient.DownloadToAsync(writableStream, cancellationToken: cancellationToken);

                // Reset the stream position to the beginning
                writableStream.Position = 0;

                return writableStream; // Return the writable stream
            }
            catch (RequestFailedException ex) when (ex.Status >= 400 && ex.Status < 500)
            {
                // Log client-side errors (HTTP 4xx)
                _logger.LogWarning(ex, "Error while trying to download Blob '{BlobName}': {Message}", blobName, ex.Message);
                throw new Exception($"Failed to download Blob '{blobName}'. Client error: {ex.Message}", ex);
            }
            catch (RequestFailedException ex)
            {
                // Log server-side errors or unexpected failures
                _logger.LogError(ex, "Unexpected error while trying to download Blob '{BlobName}': {Message}", blobName, ex.Message);
                throw;
            }
        }

        public async Task UploadFileAsync(Stream file, string containerName, string fileName, CancellationToken cancellationToken)
        {
            try
            {
                // Configure BlobClient options with retry and timeout
                var blobClientOptions = new BlobClientOptions
                {
                    Retry =
            {
                MaxRetries = _settings.MaxRetries,
                NetworkTimeout = TimeSpan.FromMinutes(_settings.NetworkTimeout)
            }
                };

                // Create container client and ensure container exists
                var blobContainerClient = new BlobContainerClient(_connectionString, containerName, blobClientOptions);
                await blobContainerClient.CreateIfNotExistsAsync(cancellationToken: cancellationToken);

                _logger.LogInformation("Container '{ContainerName}' is ready.", containerName);

                // Create blob client for the target file
                var blobClient = blobContainerClient.GetBlobClient(fileName);

                // Prepare stream for upload (buffer if non-seekable)
                Stream uploadStream;
                if (file.CanSeek)
                {
                    file.Position = 0;
                    uploadStream = file;
                }
                else
                {
                    _logger.LogInformation("Buffering non-seekable stream for upload...");
                    var buffered = new MemoryStream();
                    await file.CopyToAsync(buffered, cancellationToken);
                    buffered.Position = 0;
                    uploadStream = buffered;
                }

                // Determine MIME type based on file extension
                var contentType = Path.GetExtension(fileName).ToLowerInvariant() switch
                {
                    ".pptx" => "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    ".ppt" => "application/vnd.ms-powerpoint",
                    ".pdf" => "application/pdf",
                    ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    _ => "application/octet-stream"
                };

                var headers = new BlobHttpHeaders
                {
                    ContentType = contentType
                };

                _logger.LogInformation("Uploading file '{FileName}' ({Size} bytes) with Content-Type '{ContentType}'...",
                    fileName, uploadStream.Length, contentType);

                // Upload stream with headers and cancellation token
                await blobClient.UploadAsync(
                    uploadStream,
                    new BlobUploadOptions { HttpHeaders = headers },
                    cancellationToken);

                _logger.LogInformation("File '{FileName}' uploaded successfully to blob '{BlobUri}'.", fileName, blobClient.Uri);
            }
            catch (RequestFailedException ex) when (ex.Status >= 400 && ex.Status < 500)
            {
                _logger.LogWarning(ex, "Client error while uploading '{FileName}': {Message}", fileName, ex.Message);
                throw new Exception($"Upload failed: {ex.Message}", ex);
            }
            catch (RequestFailedException ex)
            {
                _logger.LogError(ex, "Server error while uploading '{FileName}': {Message}", fileName, ex.Message);
                throw;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Unexpected error while uploading '{FileName}': {Message}", fileName, ex.Message);
                throw;
            }
        }

        public async Task<bool> DoesBlobExistAsync(string containerName, string blobName,CancellationToken cancellationToken)
        {
            try
            {
                // 1. Create a client for the Blob Storage service
                BlobServiceClient blobServiceClient = new (_connectionString);

                // 2. Get the client for the Blob Container
                BlobContainerClient containerClient = blobServiceClient.GetBlobContainerClient(containerName);

                var blobClient = containerClient.GetBlobClient(blobName);
                return await blobClient.ExistsAsync(cancellationToken: cancellationToken);
            }
            catch (RequestFailedException ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return false;
            }
        }
    }
}

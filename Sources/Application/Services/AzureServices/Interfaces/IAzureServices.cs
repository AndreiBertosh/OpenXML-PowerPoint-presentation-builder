namespace Application.Services.AzureServices.Interfaces
{
    public interface IAzureServices
    {
        Task<Stream> GetBlobStreamAsync(string container, string blobName, CancellationToken cancellationToken);

        Task<Stream> GetWritableBlobStreamAsync(string containerName, string blobName, CancellationToken cancellationToken);

        Task UploadFileAsync(Stream file, string containerName, string fileName, CancellationToken cancellationToken);

        Task<IEnumerable<string>> GetBlobsAsync(string containerName, CancellationToken cancellationToken);

        Task<string> GetDownloadLinkAsync(string containerName, string blobName, CancellationToken cancellationToken);

        Task<bool> DeleteBlobAsync(string containerName, string blobName, CancellationToken cancellationToken);

        Task<bool> DoesBlobExistAsync(string containerName, string blobName, CancellationToken cancellationToken);

    }
}

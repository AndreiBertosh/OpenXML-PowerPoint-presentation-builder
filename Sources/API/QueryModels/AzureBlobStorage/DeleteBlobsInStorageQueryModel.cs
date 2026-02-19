namespace API.QueryModels.AzureBlobStorage
{
    public record DeleteBlobsInStorageQueryModel(
        string ContainerName,
        string[] Blobs
        );
}

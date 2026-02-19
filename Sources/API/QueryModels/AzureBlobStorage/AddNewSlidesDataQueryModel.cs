namespace API.QueryModels.AzureBlobStorage
{
    public record AddNewSlidesDataQueryModel(
        string TemplateBlobName,
        string DestinationBlobName,
        NewSlideDataQueryModel[] SlidesData
        );
}

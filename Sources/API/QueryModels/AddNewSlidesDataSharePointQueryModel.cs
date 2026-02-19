namespace API.QueryModels
{
    public record AddNewSlidesDataSharePointQueryModel(
        string TemplatePresentationName,
        string DestinationPresentationName,
        NewSlideDataQueryModel[] SlidesData
        );
}

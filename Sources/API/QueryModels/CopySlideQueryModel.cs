namespace API.QueryModels
{
    public record CopySlideQueryModel(
        string SourcePresentationName,
        string TemplatePresentationName,
        string DestinationPresentationName,
        int[] SlideIndexes,
        string CommentMessage
        );
}

namespace API.QueryModels
{
    public record CreatePresentationBySlideIndexesQueryModel(
        string SourcePresentationName,
        string DestinationPresentationName,
        int[] SlideIndexes,
        string CommentMessage
        );
}

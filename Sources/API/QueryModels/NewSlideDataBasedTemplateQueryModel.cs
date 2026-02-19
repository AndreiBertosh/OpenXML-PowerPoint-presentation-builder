namespace API.QueryModels
{
    public record NewSlideDataBasedTemplateQueryModel(
        int SlideIndex,
        string TitleLine1,
        string TitleLine2,
        string TitleLine3,
        string[] TextBlock,
        string ColumnBlockHeader,
        string ColumnHeader1,
        string ColumnText1,
        string ColumnHeader2,
        string ColumnText2,
        string ColumnHeader3,
        string ColumnText3
        );
}

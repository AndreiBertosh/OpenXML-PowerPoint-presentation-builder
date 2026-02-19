namespace API.Endpoints.PresentationParser.Models
{
    public record SlideDataViewModel(
        uint SlideId,
        int SlideIndex,
        string SlideRelationshipId,
        SectionDataViewModel Section,
        SlideTitleDataViewModel SlideTitle,
        string[] Texts,
        string[] Notes);
}

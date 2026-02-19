namespace Application.Services.PresentationParser.Models
{
    public record SlideDataDomainModel(
        uint SlideId,
        int SlideIndex,
        string SlideRelationshipId,
        SectionDataDomainModel Section,
        SlideTitleDataDomainModel SlideTitle,
        string[] Texts,
        string[] Notes);
}

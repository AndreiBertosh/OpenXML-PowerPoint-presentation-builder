using Application.Services.PresentationParser.Models;

using DocumentFormat.OpenXml.Packaging;

namespace Application.Services.PresentationParser.Interfaces
{
    public interface IPresentationParserServices
    {
        List<SlideDataDomainModel> GetAllPresentationData(PresentationDocument presentationDocument);

        //List<SlideMasterDataDomainModel> GetAllLayoutsPresentationData(PresentationDocument presentationDocument);

        IReadOnlyList<SlideMasterInfoDomainModel> AnalyzePresentationLayouts(PresentationDocument presentationDocument);
    }
}

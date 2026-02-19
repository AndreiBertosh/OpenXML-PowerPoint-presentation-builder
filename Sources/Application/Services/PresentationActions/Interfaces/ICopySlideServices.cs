using DocumentFormat.OpenXml.Packaging;

using Microsoft.Extensions.Logging;

namespace Application.Services.PresentationActions.Interfaces
{
    public interface ICopySlideServices
    {
        Task<IEnumerable<string>> CopySlides(PresentationDocument sourceDocument, PresentationDocument destDocument, int[] copiedSlidePositions, string comment, ILogger logger);

        Task<SlidePart> CopySlideWithLayout(PresentationDocument sourceDocument, PresentationDocument targetDocument, int sourceSlideIndex, string comment, int? targetPosition = null);
    }
}

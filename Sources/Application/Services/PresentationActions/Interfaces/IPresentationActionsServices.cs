using DocumentFormat.OpenXml.Packaging;

using Microsoft.Extensions.Logging;

namespace Application.Services.PresentationActions.Interfaces
{
    public interface IPresentationActionsServices
    {
        Task<IEnumerable<string>> NewPresentationBySlideIndexes(PresentationDocument presentationDocument, List<int> indexesForSave, string message, ILogger logger);
    }
}

using Application.Services.PresentationActions.Models;

using DocumentFormat.OpenXml.Packaging;

using Microsoft.Extensions.Logging;

namespace Application.Services.PresentationActions.Interfaces
{
    public interface IAddNewSlideToPresentationServices
    {
        IEnumerable<string>? AddNewSlideByLayout(PresentationDocument presentationDocument, NewSlideData slideData, ILogger logger);

        IEnumerable<string>? AddNewSlidesByLayout(PresentationDocument presentationDocument, NewSlideData[] slidesData, ILogger logger);

        //IEnumerable<string>? AddNewSlidesByLayout(PresentationDocument sourceDocument, PresentationDocument targetDocument, NewSlideBasedTemplateData[] slidesData, string comment, ILogger logger);

    }
}

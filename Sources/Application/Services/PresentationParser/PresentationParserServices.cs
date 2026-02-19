using Application.Services.Common;
using Application.Services.PresentationParser.Interfaces;
using Application.Services.PresentationParser.Models;
using Application.Services.PresentationParser.SlideLayoutsParser.Parsers;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.PowerPoint;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

using Drawing = DocumentFormat.OpenXml.Drawing;

namespace Application.Services.PresentationParser
{
    public class PresentationParserServices : IPresentationParserServices
    {
        public List<SlideDataDomainModel> GetAllPresentationData(PresentationDocument presentationDocument)
        {
            ArgumentNullException.ThrowIfNull(presentationDocument);

            List<SlideDataDomainModel> slides = [];

            int slideIdsCount = PresentationCommonServices.SlideIdsCount(presentationDocument);

            if (slideIdsCount > 0)
            {
                PresentationPart? presentationPart = presentationDocument?.PresentationPart;
                Presentation? presentation = presentationPart?.Presentation;
                SlideIdList? slideIdList = presentation?.SlideIdList;

                if (slideIdList != null)
                {
                    for (int index = 0; index < slideIdsCount; index++)
                    {
                        SlideId? slideId = slideIdList?.ChildElements[index] as SlideId;
                        string? relationshipId = slideId?.RelationshipId?.Value;

                        if (relationshipId is null)
                        {
                            continue;
                        }

                        if (presentationPart?.GetPartById(relationshipId) is SlidePart slidePart)
                        {
                            SlideDataDomainModel slideModel = new SlideDataDomainModel(
                                SlideId: slideId.Id.Value,
                                SlideIndex: index + 1,
                                SlideRelationshipId: relationshipId,
                                Section: new SectionDataDomainModel(string.Empty, string.Empty),
                                SlideTitle: GetSlideTitle(slidePart),
                                Texts: GetAllTextDataFromSlide(slidePart),
                                Notes: GetAllTextDataFromNoteSlide(slidePart)
                            );

                            slides.Add(slideModel);
                        }
                    }
                }
            }

            return AddSectionDataToSlides(presentationDocument, slides);
        }

        private static SlideTitleDataDomainModel GetSlideTitle(SlidePart slidePart)
        {
            IEnumerable<Shape> shapes = slidePart.Slide.Descendants<Shape>();

            string title = string.Empty;
            string subTitle = string.Empty;
            List<string> titles = new();
            List<string> subtitles = new();

            foreach (Shape shape in shapes)
            {
                PlaceholderShape? placeholder = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape;
                if (placeholder != null && placeholder.Type != null)
                {
                    if (placeholder.Type.Value == PlaceholderValues.Title)
                    {
                        titles.AddRange(shape.TextBody?.Descendants<Drawing.Paragraph>().Select(static t => t.InnerText).ToList());
                    }

                    if (placeholder.Type.Value == PlaceholderValues.SubTitle)
                    {
                        subtitles.AddRange(shape.TextBody?.Descendants<Drawing.Paragraph>().Select(static t => t.InnerText).ToList());
                    }
                }
            }

            if (titles.Count > 0)
            {
                title = string.Join(" ", titles);
            }

            if (subtitles.Count > 0)
            {
                subTitle = string.Join(" ", subtitles);
            }

            return new SlideTitleDataDomainModel(title, subTitle);
        }

        private static string[] GetAllTextDataFromNoteSlide(SlidePart slidePart)
        {

            var notesPart = slidePart.NotesSlidePart;

            if (notesPart != null)
            {
                var notesSlide = notesPart.NotesSlide;

                if (notesSlide != null)
                {
                    var shape = notesSlide.Descendants<Shape>()
                        .FirstOrDefault(s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type == PlaceholderValues.Body);

                    var paragraphs = shape.Descendants<Drawing.Paragraph>().ToList();

                    if (paragraphs.Count > 0)
                    {
                        return paragraphs
                        .SelectMany(paragraph => paragraph.Descendants<Drawing.Text>())
                        .Select(text => text.Text)
                        .ToArray();
                    }
                    else
                    {
                        return Array.Empty<string>();
                    }
                }
            }

            return Array.Empty<string>();
        }

        private static string[] GetAllTextDataFromSlide(SlidePart slidePart)
        {
            LinkedList<string> slideTexts = new();

            if (slidePart != null)
            {
                foreach (Drawing.Paragraph paragraph in slidePart.Slide.Descendants<Drawing.Paragraph>())
                {
                    if (!string.IsNullOrWhiteSpace(paragraph.InnerText))
                    {
                        slideTexts.AddLast(paragraph.InnerText);
                    }
                }
            }

            return slideTexts.ToArray();
        }

        private static List<SlideDataDomainModel> AddSectionDataToSlides(PresentationDocument presentationDocument, List<SlideDataDomainModel> slidesData)
        {
            List<SectionSlideIdDomainModel> sectionSlideIds = GetSectionSlideIds(presentationDocument);

            List<SlideDataDomainModel> resultList = slidesData
                .Select(slideData => slideData with
                {
                    Section = new SectionDataDomainModel(
                        sectionSlideIds
                            .Where(section => section.SlideId == slideData.SlideId)
                            .Select(section => section.SectionId)
                            .FirstOrDefault(),
                        sectionSlideIds
                            .Where(section => section.SlideId == slideData.SlideId)
                            .Select(section => section.SectionName)
                            .FirstOrDefault())
                })
                .ToList();

            return resultList;
        }

        private static List<SectionSlideIdDomainModel> GetSectionSlideIds(PresentationDocument presentationDocument)
        {
            List<SectionSlideIdDomainModel> sectionSlideIds = new();

            PresentationPart? presentationPart = presentationDocument.PresentationPart;
            Presentation? presentation = presentationPart?.Presentation;
            PresentationExtensionList? presentationExtensionList = presentation?.PresentationExtensionList;

            SectionList sectionList = new();

            if (presentationExtensionList != null)
            {
                foreach (PresentationExtension presentationExtension in presentationExtensionList)
                {
                    sectionList = presentationExtension.GetFirstChild<SectionList>();

                    if (sectionList != null)
                    {
                        break;
                    }
                }

                if (sectionList != null)
                {
                    foreach (Section section in sectionList.ChildElements)
                    {
                        SectionSlideIdList? sectionSlideIdList = section.SectionSlideIdList;
                        List<SectionSlideIdDomainModel> temptempSectionSlideIds =
                            sectionSlideIdList?.Select(slide =>
                                new SectionSlideIdDomainModel(
                                    section.Id,
                                    section.Name,
                                    (slide as SectionSlideIdListEntry).Id.Value
                                )
                            ).ToList();

                        if (temptempSectionSlideIds != null)
                        {
                            sectionSlideIds.AddRange(temptempSectionSlideIds);
                        }
                    }
                }
            }

            return sectionSlideIds;
        }

        public IReadOnlyList<SlideMasterInfoDomainModel> AnalyzePresentationLayouts(PresentationDocument presentationDocument)
        {
            // Guard
            var presPart = presentationDocument.PresentationPart!;
            var result = new List<SlideMasterInfoDomainModel>();

            // Iterate over all slide masters and delegate to SlideMasterParser
            foreach (var masterPart in presPart.SlideMasterParts)
            {
                var masterModel = SlideMasterParser.Parse(masterPart, presPart);
                result.Add(masterModel);
            }

            return result;
        }
    }
}

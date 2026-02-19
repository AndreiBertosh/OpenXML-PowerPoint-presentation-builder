using Application.Services.Common;
using Application.Services.PresentationActions.Interfaces;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.PowerPoint;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

using Microsoft.Extensions.Logging;

namespace Application.Services.PresentationActions
{
    public class PresentationActionsServices : IPresentationActionsServices
    {
        public Task<IEnumerable<string>> NewPresentationBySlideIndexes(PresentationDocument presentationDocument, List<int> indexesForSave, string message, ILogger logger)
        {
            int slideIdsCount = PresentationCommonServices.SlideIdsCount(presentationDocument);

            PresentationPart? presentationPart = presentationDocument.PresentationPart;
            Presentation? presentation = presentationPart?.Presentation;
            SlideIdList? slideIdList = presentation?.SlideIdList;

            if (slideIdList != null)
            {
                OpenXmlElementList slideIds = slideIdList.ChildElements;

                for (int index = slideIdsCount - 1; index >= 0; index--)
                {
                    if (!indexesForSave.Contains(index + 1))
                    {
                        SlideId? slideId = slideIdList.ChildElements[index] as SlideId;
                        string slideRelationshipId = slideId?.RelationshipId;

                        if (slideRelationshipId != null)
                        {
                            slideIdList?.RemoveChild(slideId);
                        }

                        if (presentation!.CustomShowList is not null)
                        {
                            // Iterate through the list of custom shows.
                            foreach (var customShow in presentation.CustomShowList.Elements<CustomShow>())
                            {
                                if (customShow.SlideList is not null)
                                {
                                    // Declare a link list of slide list entries.
                                    LinkedList<SlideListEntry> slideListEntries = new LinkedList<SlideListEntry>();
                                    foreach (SlideListEntry slideListEntry in customShow.SlideList.Elements())
                                    {
                                        // Find the slide reference to remove from the custom show.
                                        if (slideListEntry.Id is not null && slideListEntry.Id == slideRelationshipId)
                                        {
                                            slideListEntries.AddLast(slideListEntry);
                                        }
                                    }

                                    // Remove all references to the slide from the custom show.
                                    foreach (SlideListEntry slideListEntry in slideListEntries)
                                    {
                                        customShow.SlideList.RemoveChild(slideListEntry);
                                    }
                                }
                            }
                        }

                        // Get the slide part for the specified slide.
                        SlidePart slidePart = (SlidePart)presentationPart!.GetPartById(slideRelationshipId);

                        // Remove the slide part.
                        presentationPart.DeletePart(slidePart);
                    }
                }

                if (!string.IsNullOrEmpty(message))
                {
                    foreach (SlideId slideId in slideIdList)
                    {
                        var slidePart = presentationPart.GetPartById(slideId.RelationshipId) as SlidePart;
                        PresentationCommonServices.AddNoteToSlide(presentationPart, slidePart, message);
                    }
                }
            }

            presentationDocument.Save();
            ClearSections(presentation);

            PresentationCommonServices.ClearAllComments(presentation);

            // Validate the document
            var errors = PresentationCommonServices.ValidateDocument(presentationDocument, logger);

            return Task.FromResult(errors);
        }

        private void ClearSections(Presentation? presentation)
        {
            if (presentation != null)
            {
                PresentationExtensionList? presentationExtensionList = presentation.PresentationExtensionList;

                SectionList? sectionList = null;

                if (presentationExtensionList != null)
                {
                    foreach (PresentationExtension presentationExtension in presentationExtensionList.Cast<PresentationExtension>())
                    {
                        sectionList = presentationExtension.GetFirstChild<SectionList>();

                        if (sectionList != null)
                        {
                            break;
                        }
                    }

                    if (sectionList != null)
                    {
                        Stack<Section> sections = new();

                        foreach (Section section in sectionList)
                        {
                            sections.Push(section);
                        }

                        List<uint> slideIds = GetPresentationSlideIds(presentation);

                        foreach (Section section in sections)
                        {
                            SectionSlideIdList sectionSlideIdList = section.SectionSlideIdList;
                            List<SectionSlideIdListEntry> sectionSlideIdListEntries = sectionSlideIdList?.Select(slide => slide as SectionSlideIdListEntry).ToList();

                            if (sectionSlideIdListEntries != null)
                            {
                                foreach (SectionSlideIdListEntry sectionSlideIdEntry in sectionSlideIdListEntries)
                                {
                                    if (!slideIds.Contains(sectionSlideIdEntry.Id.Value))
                                    {
                                        section.SectionSlideIdList.RemoveChild(sectionSlideIdEntry);
                                    }
                                }
                            }

                            if (section.SectionSlideIdList == null || section.SectionSlideIdList.Count() == 0)
                            {
                                sectionList.RemoveChild(section);
                            }
                        }
                    }
                }
            }
        }

        private List<uint> GetPresentationSlideIds(Presentation presentation)
        {
            List<uint> presentationSlideIds = new();

            SlideIdList? slideIdList = presentation.SlideIdList;
            if (slideIdList != null)
            {
                foreach (SlideId slideId in slideIdList)
                {
                    presentationSlideIds.Add(slideId.Id);
                }
            }

            return presentationSlideIds;
        }
    }
}

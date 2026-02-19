using Application.Services.Common;
using Application.Services.PresentationActions.Interfaces;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

using Microsoft.CodeAnalysis;
using Microsoft.Extensions.Logging;

using Drawing = DocumentFormat.OpenXml.Drawing;

namespace Application.Services.PresentationActions
{
    public class CopySlideServices(ISlideMasterCopyService slideMasterCopyService) : ICopySlideServices
    {
        private readonly ISlideMasterCopyService _slideMasterCopyService = slideMasterCopyService;

        public async Task<IEnumerable<string>> CopySlides(
            PresentationDocument sourceDocument,
            PresentationDocument targetDocument,
            int[] sourceSlideIndexes,
            string comment,
            ILogger logger)
        {
            var sourcePresPart = sourceDocument.PresentationPart;
            var targetPresPart = targetDocument.PresentationPart;

            // Caches for masters and layouts (to prevent duplication)
            var masterCache = new Dictionary<string, SlideMasterPart>(StringComparer.OrdinalIgnoreCase);
            var layoutCache = new Dictionary<(string Theme, string Layout), SlideLayoutPart>();

            foreach (var index in sourceSlideIndexes)
            {
                CopySlideWithLayout(
                    sourceDocument,
                    targetDocument,
                    index - 1,
                    comment,
                    masterCache,
                    layoutCache
                );
            }

            targetPresPart.Presentation.Save();
            logger.LogInformation("Presentation saved. Validating…");

            var result = await Task.Run(() => PresentationCommonServices.ValidateDocument(targetDocument, logger));
            logger.LogInformation("Validation complete. {ErrorCount} issues found.", result.Count());

            return result;
        }


        public Task<SlidePart> CopySlideWithLayout(
            PresentationDocument sourceDoc,
            PresentationDocument targetDoc,
            int sourceSlideIndex,
            string comment,
            Dictionary<string, SlideMasterPart> masterCache,
            Dictionary<(string, string), SlideLayoutPart> layoutCache,
            int? targetPosition = null)
        {
            // Resolve parts
            PresentationPart sourcePresPart = sourceDoc.PresentationPart
                ?? throw new InvalidOperationException("Source presentation has no PresentationPart.");

            PresentationPart targetPresPart = targetDoc.PresentationPart
                ?? throw new InvalidOperationException("Target presentation has no PresentationPart.");

            // Resolve the source slide
            SlidePart sourceSlidePart = GetSlidePartByIndex(sourcePresPart, sourceSlideIndex)
                ?? throw new ArgumentException("Source slide not found.");

            // Extract identifiers
            string sourceThemeName = GetSlideThemeName(sourceSlidePart);
            string sourceLayoutName = GetLayoutName(sourceSlidePart.SlideLayoutPart);
            string sourceLayoutType = GetLayoutType(sourceSlidePart.SlideLayoutPart);

            // 1. Ensure SlideMaster exists in target (MAIN REQUIREMENT)
            SlideMasterPart targetMasterPart = null;

            if (!string.IsNullOrWhiteSpace(sourceThemeName) &&
                !string.Equals(sourceThemeName, "Unknown", StringComparison.OrdinalIgnoreCase))
            {
                // First check master cache
                if (!masterCache.TryGetValue(sourceThemeName, out targetMasterPart))
                {
                    // Try find in target
                    targetMasterPart = targetPresPart.SlideMasterParts
                        .FirstOrDefault(slideMasterPart =>
                            string.Equals(slideMasterPart.ThemePart?.Theme?.Name?.Value,
                                          sourceThemeName,
                                          StringComparison.OrdinalIgnoreCase));

                    // If missing → copy via injected service
                    if (targetMasterPart == null)
                    {
                        // Copy SlideMaster and get its RelationshipId in the target
                        string masterRelId = _slideMasterCopyService.CopySlideMasterByThemeName(
                            sourcePresPart,
                            targetPresPart,
                            sourceThemeName,
                            skipIfExistsInTarget: true).Result;

                        // Now fetch the newly-copied SlideMasterPart by RelationshipId
                        targetMasterPart = (SlideMasterPart)targetPresPart.GetPartById(masterRelId);

                        if (targetMasterPart == null)
                        {
                            throw new InvalidOperationException(
                                $"SlideMaster copy succeeded but cannot resolve the copied master by rId '{masterRelId}'.");
                        }
                    }

                    masterCache[sourceThemeName] = targetMasterPart ?? throw new InvalidOperationException(
                            $"Failed to ensure SlideMaster with theme '{sourceThemeName}' exists in target.");
                }
            }

            // 2. Create new slide
            SlidePart newSlidePart = targetPresPart.AddNewPart<SlidePart>();

            // Copy slide XML + relationships
            CopySlideContent(
                sourceSlidePart,
                newSlidePart,
                targetPresPart,
                sourceThemeName,
                sourceLayoutName,
                sourceLayoutType);

            // 3. Ensure layout exists & bind it
            SlideLayoutPart sourceLayoutPart = sourceSlidePart.SlideLayoutPart;

            SlideLayoutPart ensuredLayout = EnsureLayoutExistsWithMaster(
                targetPresPart,
                sourceLayoutPart,
                masterCache,
                layoutCache);

            if (!newSlidePart.Parts.Any(p => p.OpenXmlPart == ensuredLayout))
                newSlidePart.AddPart(ensuredLayout);

            // 4. Notes
            CopyNotesWithContent(sourceSlidePart, newSlidePart);
            PresentationCommonServices.AddNoteToSlide(targetPresPart, newSlidePart, comment);

            // 5. Insert into slide list
            AddSlideToPresentation(targetPresPart, newSlidePart, targetPosition);

            targetPresPart.Presentation.Save();

            return Task.FromResult(newSlidePart);
        }

        private static SlideLayoutPart EnsureLayoutExistsWithMaster(
            PresentationPart targetPresPart,
            SlideLayoutPart sourceLayoutPart,
            Dictionary<string, SlideMasterPart> masterCache,
            Dictionary<(string, string), SlideLayoutPart> layoutCache)
        {
            var sourceMasterPart = sourceLayoutPart.SlideMasterPart;
            string themeName = sourceMasterPart?.ThemePart?.Theme?.Name?.Value ?? "";
            string layoutName = GetLayoutName(sourceLayoutPart);

            // 1. Master cache
            if (!masterCache.TryGetValue(themeName, out SlideMasterPart targetMasterPart))
            {
                // Find SlideMaster in target by theme name
                targetMasterPart = targetPresPart.SlideMasterParts
                    .FirstOrDefault(mp =>
                        mp.ThemePart?.Theme?.Name?.Value != null &&
                        mp.ThemePart.Theme.Name.Value.Equals(themeName, StringComparison.OrdinalIgnoreCase));

                // If not found, copy master (with all layouts and theme part)
                if (targetMasterPart == null)
                {
                    targetMasterPart = targetPresPart.AddNewPart<SlideMasterPart>();
                    targetMasterPart.FeedData(sourceMasterPart.GetStream());

                    if (sourceMasterPart.ThemePart != null && targetMasterPart.ThemePart == null)
                    {
                        var newThemePart = targetMasterPart.AddNewPart<ThemePart>();
                        newThemePart.FeedData(sourceMasterPart.ThemePart.GetStream());
                    }
                }
                masterCache[themeName] = targetMasterPart;
            }

            // 2. Layout cache
            var layoutKey = (themeName, layoutName);
            if (layoutCache.TryGetValue(layoutKey, out var cachedLayout))
            {
                return cachedLayout;
            }

            // 3. Find or copy layout in master
            SlideLayoutPart matchingLayout = targetMasterPart.SlideLayoutParts
                .FirstOrDefault(lp => GetLayoutName(lp).Equals(layoutName, StringComparison.OrdinalIgnoreCase));

            if (matchingLayout == null)
            {
                matchingLayout = targetMasterPart.AddNewPart<SlideLayoutPart>();
                matchingLayout.FeedData(sourceLayoutPart.GetStream());
                foreach (var img in sourceLayoutPart.ImageParts)
                {
                    var newImg = matchingLayout.AddImagePart(img.ContentType);
                    using var imgStream = img.GetStream();
                    newImg.FeedData(imgStream);
                }
            }

            layoutCache[layoutKey] = matchingLayout;
            return matchingLayout;
        }

        public Task<SlidePart> CopySlideWithLayout(PresentationDocument sourceDoc, PresentationDocument targetDoc, int sourceSlideIndex, string comment, int? targetPosition = null)
        {
            PresentationPart sourcePresPart = sourceDoc.PresentationPart;
            PresentationPart targetPresPart = targetDoc.PresentationPart;

            SlidePart sourceSlidePart = GetSlidePartByIndex(sourcePresPart, sourceSlideIndex) ?? throw new ArgumentException("Source slide not found");

            // Get source theme information and layout details
            string sourceThemeName = GetSlideThemeName(sourceSlidePart);
            string sourceLayoutName = GetLayoutName(sourceSlidePart.SlideLayoutPart);
            string sourceLayoutType = GetLayoutType(sourceSlidePart.SlideLayoutPart);

            // Create new slide part in target presentation
            SlidePart newSlidePart = targetPresPart.AddNewPart<SlidePart>();

            // Copy all slide content including relationships
            CopySlideContent(sourceSlidePart, newSlidePart, targetPresPart, sourceThemeName, sourceLayoutName, sourceLayoutType);

            // Copy notes with content from source slide
            CopyNotesWithContent(sourceSlidePart, newSlidePart);

            // Add notes with content from comment value
            PresentationCommonServices.AddNoteToSlide(targetPresPart, newSlidePart, comment);

            // Add to presentation
            AddSlideToPresentation(targetPresPart, newSlidePart, targetPosition);

            // Save changes to ensure relationships are persisted
            targetPresPart.Presentation.Save();

            return Task.FromResult(newSlidePart);
            //return Task.CompletedTask;
        }

        private static string GetSlideThemeName(SlidePart slidePart)
        {
            try
            {
                var layoutPart = slidePart.SlideLayoutPart;
                if (layoutPart != null)
                {
                    var masterPart = layoutPart.SlideMasterPart;
                    if (masterPart != null && masterPart.ThemePart != null && masterPart.ThemePart.Theme != null)
                    {
                        return GetThemeName(masterPart);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to get slide theme name: {ex.Message}");
            }
            return "Unknown";
        }

        private static string GetThemeName(SlideMasterPart slideMasterPart)
        {
            try
            {
                if (slideMasterPart?.ThemePart?.Theme?.Name?.Value != null)
                {
                    return slideMasterPart.ThemePart.Theme.Name.Value;
                }
            }
            catch
            {
                // Ignore errors
            }
            return "Unknown";
        }

        private static string GetLayoutName(SlideLayoutPart layoutPart)
        {
            if (!string.IsNullOrEmpty(layoutPart.SlideLayout.CommonSlideData.Name))
            {
                return layoutPart.SlideLayout.CommonSlideData.Name;
            }

            try
            {
                // Use URI-based name - this is the most reliable way to identify layouts
                if (layoutPart?.Uri != null)
                {
                    string fileName = Path.GetFileNameWithoutExtension(layoutPart.Uri.ToString());

                    // Try to extract meaningful name from URI
                    if (fileName.StartsWith("slideLayout"))
                    {
                        // For standard layouts like "slideLayout1.xml", "slideLayout2.xml"
                        return $"Layout_{fileName.Substring(11)}";
                    }
                    else
                    {
                        return fileName;
                    }
                }
            }
            catch
            {
                // Ignore errors
            }
            return string.Empty;
        }

        private static string GetLayoutType(SlideLayoutPart layoutPart)
        {
            try
            {
                if (layoutPart?.SlideLayout?.Type?.Value != null)
                {
                    return layoutPart.SlideLayout.Type.Value.ToString();
                }
            }
            catch
            {
                // Ignore errors
            }
            return string.Empty;
        }

        private static void CopyNotesWithContent(SlidePart sourceSlidePart, SlidePart targetSlidePart)
        {
            try
            {
                // Check if source slide has notes
                NotesSlidePart sourceNotesPart = sourceSlidePart.NotesSlidePart;
                if (sourceNotesPart == null)
                {
                    // If no notes in source, create empty notes
                    AddEmptyNotesSlide(targetSlidePart);
                    return;
                }

                // Create notes part in target slide
                NotesSlidePart targetNotesPart = targetSlidePart.AddNewPart<NotesSlidePart>();

                // Copy the notes slide structure from source
                if (sourceNotesPart.NotesSlide != null)
                {
                    // Clone the entire notes slide structure
                    targetNotesPart.NotesSlide = (NotesSlide)sourceNotesPart.NotesSlide.CloneNode(true);

                    // Copy notes relationships (images, etc.)
                    CopyNotesRelationships(sourceNotesPart, targetNotesPart);
                }
                else
                {
                    // Fallback: create empty notes if source structure is invalid
                    AddEmptyNotesSlide(targetSlidePart);
                }
            }
            catch (Exception ex)
            {
                // Log but don't fail if notes creation fails
                System.Diagnostics.Debug.WriteLine($"Failed to copy notes: {ex.Message}");
                // Fallback to empty notes
                AddEmptyNotesSlide(targetSlidePart);
            }
        }

        private static void CopyNotesRelationships(NotesSlidePart sourceNotesPart, NotesSlidePart targetNotesPart)
        {
            try
            {
                // Copy image relationships from notes
                foreach (ImagePart imagePart in sourceNotesPart.ImageParts)
                {
                    ImagePart newImagePart = targetNotesPart.AddImagePart(imagePart.ContentType);
                    using (Stream stream = imagePart.GetStream())
                    {
                        newImagePart.FeedData(stream);
                    }
                }

                // Copy external relationships from notes
                foreach (var relationship in sourceNotesPart.ExternalRelationships)
                {
                    targetNotesPart.AddExternalRelationship(relationship.RelationshipType, relationship.Uri, relationship.Id);
                }

                // Copy hyperlink relationships from notes
                foreach (var relationship in sourceNotesPart.HyperlinkRelationships)
                {
                    targetNotesPart.AddHyperlinkRelationship(relationship.Uri, relationship.IsExternal, relationship.Id);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to copy notes relationships: {ex.Message}");
            }
        }

        private static void AddEmptyNotesSlide(SlidePart slidePart)
        {
            try
            {
                // Create empty notes part
                NotesSlidePart notesPart = slidePart.AddNewPart<NotesSlidePart>();

                // Create valid notes slide structure
                NotesSlide notesSlide = new NotesSlide();

                // Add required CommonSlideData with properly structured ShapeTree
                CommonSlideData commonSlideData = new CommonSlideData();
                ShapeTree shapeTree = CreateValidShapeTree();
                commonSlideData.Append(shapeTree);
                notesSlide.Append(commonSlideData);

                // Add ColorMapOverride
                notesSlide.Append(new ColorMapOverride(
                    new Drawing.MasterColorMapping()
                ));

                notesPart.NotesSlide = notesSlide;
            }
            catch (Exception ex)
            {
                // Log but don't fail if notes creation fails
                System.Diagnostics.Debug.WriteLine($"Failed to create empty notes: {ex.Message}");
            }
        }

        private static ShapeTree CreateValidShapeTree()
        {
            // Create a valid ShapeTree with required non-visual properties
            ShapeTree shapeTree = new ShapeTree();

            // Add required NonVisualGroupShapeProperties (nvGrpSpPr)
            NonVisualGroupShapeProperties nonVisualProperties = new (
                new NonVisualDrawingProperties() { Id = 1U, Name = "Notes Placeholder" },
                new NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()
            );

            shapeTree.Append(nonVisualProperties);

            // Add GroupShapeProperties (optional but good practice)
            shapeTree.Append(new GroupShapeProperties());

            return shapeTree;
        }

        private static void CopySlideContent(SlidePart sourceSlidePart, SlidePart targetSlidePart, PresentationPart targetPresPart, string sourceThemeName, string sourceLayoutName, string sourceLayoutType)
        {
            // Clone the slide XML first
            targetSlidePart.Slide = (Slide)sourceSlidePart.Slide.CloneNode(true);

            // Copy all relationships before modifying the XML
            CopySlideRelationships(sourceSlidePart, targetSlidePart, targetPresPart);

            // Update layout reference after copying relationships
            UpdateSlideLayoutReference(sourceSlidePart, targetSlidePart, targetPresPart, sourceThemeName, sourceLayoutName, sourceLayoutType);
        }

        private static void CopySlideRelationships(SlidePart sourceSlidePart, SlidePart targetSlidePart, PresentationPart targetPresPart)
        {
            // Copy image relationships
            foreach (ImagePart imagePart in sourceSlidePart.ImageParts)
            {
                string sourceRelId = sourceSlidePart.GetIdOfPart(imagePart);
                ImagePart newImagePart = targetSlidePart.AddImagePart(imagePart.ContentType, sourceRelId);
                using (Stream stream = imagePart.GetStream())
                {
                    newImagePart.FeedData(stream);
                }
            }

            // Copy external relationships
            foreach (var relationship in sourceSlidePart.ExternalRelationships)
            {
                targetSlidePart.AddExternalRelationship(relationship.RelationshipType, relationship.Uri, relationship.Id);
            }

            // Copy hyperlink relationships
            foreach (var relationship in sourceSlidePart.HyperlinkRelationships)
            {
                targetSlidePart.AddHyperlinkRelationship(relationship.Uri, relationship.IsExternal, relationship.Id);
            }
        }

        private static void UpdateSlideLayoutReference(
            SlidePart sourceSlidePart,
            SlidePart targetSlidePart,
            PresentationPart targetPresPart,
            string sourceThemeName,
            string sourceLayoutName,
            string sourceLayoutType)
        {
            var sourceLayoutPart = sourceSlidePart.SlideLayoutPart;
            if (sourceLayoutPart == null) return;

            // First try to find layout in a master with the same theme name
            SlideLayoutPart targetLayoutPart = FindLayoutInMatchingTheme(
                targetPresPart,
                sourceThemeName,
                sourceLayoutName,
                sourceLayoutType,
                sourceLayoutPart);

            // If not found, fallback to the previous global algorithm
            if (targetLayoutPart == null)
            {
                targetLayoutPart = FindOrCreateMatchingLayout(
                    targetPresPart,
                    sourceLayoutPart,
                    sourceLayoutName,
                    sourceLayoutType);
            }

            if (targetLayoutPart != null)
            {
                // Correct way: just add the layout part as relationship
                if (!targetSlidePart.Parts.Any(p => p.OpenXmlPart == targetLayoutPart))
                {
                    targetSlidePart.AddPart(targetLayoutPart);
                }
            }
        }

        private static SlideLayoutPart FindLayoutInMatchingTheme(
            PresentationPart targetPresPart,
            string sourceThemeName,
            string sourceLayoutName,
            string sourceLayoutType,
            SlideLayoutPart sourceLayoutPart)
        {
            if (!string.IsNullOrEmpty(sourceThemeName) && sourceThemeName != "Unknown")
            {
                foreach (var slideMaster in targetPresPart.SlideMasterParts)
                {
                    string targetThemeName = GetThemeName(slideMaster);
                    if (string.Equals(targetThemeName, sourceThemeName, StringComparison.OrdinalIgnoreCase))
                    {
                        // 1. Try exact layout name
                        if (!string.IsNullOrEmpty(sourceLayoutName))
                        {
                            foreach (var existingLayout in slideMaster.SlideLayoutParts)
                            {
                                string existingLayoutName = GetLayoutName(existingLayout);
                                if (string.Equals(existingLayoutName, sourceLayoutName, StringComparison.OrdinalIgnoreCase))
                                {
                                    return existingLayout;
                                }
                            }
                        }

                        // 2. Try layout type
                        if (!string.IsNullOrEmpty(sourceLayoutType))
                        {
                            foreach (var existingLayout in slideMaster.SlideLayoutParts)
                            {
                                string existingLayoutType = GetLayoutType(existingLayout);
                                if (string.Equals(existingLayoutType, sourceLayoutType, StringComparison.OrdinalIgnoreCase))
                                {
                                    return existingLayout;
                                }
                            }
                        }

                        // 3. Try compatibility check
                        foreach (var existingLayout in slideMaster.SlideLayoutParts)
                        {
                            if (AreLayoutsCompatible(existingLayout, sourceLayoutPart))
                            {
                                return existingLayout;
                            }
                        }

                        // 4. Fallback: first layout from this master
                        return slideMaster.SlideLayoutParts.FirstOrDefault();
                    }
                }
            }

            return null; // If not found, return null so caller can fallback
        }

        private static SlideLayoutPart FindOrCreateMatchingLayout(
            PresentationPart targetPresPart,
            SlideLayoutPart sourceLayoutPart,
            string sourceLayoutName = "",
            string sourceLayoutType = "")
        {
            // 1. Try exact name match across all slide masters
            if (!string.IsNullOrEmpty(sourceLayoutName))
            {
                foreach (var slideMaster in targetPresPart.SlideMasterParts)
                {
                    foreach (var existingLayout in slideMaster.SlideLayoutParts)
                    {
                        string existingLayoutName = GetLayoutName(existingLayout);
                        if (string.Equals(existingLayoutName, sourceLayoutName, StringComparison.OrdinalIgnoreCase))
                        {
                            return existingLayout;
                        }
                    }
                }
            }

            // 2. Try type match
            if (!string.IsNullOrEmpty(sourceLayoutType))
            {
                foreach (var slideMaster in targetPresPart.SlideMasterParts)
                {
                    foreach (var existingLayout in slideMaster.SlideLayoutParts)
                    {
                        string existingLayoutType = GetLayoutType(existingLayout);
                        if (string.Equals(existingLayoutType, sourceLayoutType, StringComparison.OrdinalIgnoreCase))
                        {
                            return existingLayout;
                        }
                    }
                }
            }

            // 3. Try compatibility matching
            foreach (var slideMaster in targetPresPart.SlideMasterParts)
            {
                foreach (var existingLayout in slideMaster.SlideLayoutParts)
                {
                    if (AreLayoutsCompatible(existingLayout, sourceLayoutPart))
                    {
                        return existingLayout;
                    }
                }
            }

            // 4. Final fallback: use first available layout
            var firstSlideMaster = targetPresPart.SlideMasterParts.FirstOrDefault();
            if (firstSlideMaster != null)
            {
                var defaultLayout = firstSlideMaster.SlideLayoutParts.FirstOrDefault();
                if (defaultLayout != null)
                {
                    return defaultLayout;
                }
            }

            return null;
        }

        private static bool AreLayoutsCompatible(SlideLayoutPart layout1, SlideLayoutPart layout2)
        {
            try
            {
                // Compare by type
                var layout1Type = layout1.SlideLayout?.Type?.Value;
                var layout2Type = layout2.SlideLayout?.Type?.Value;

                if (layout1Type != null && layout2Type != null && layout1Type == layout2Type)
                {
                    return true;
                }

                // Compare by name
                string layout1Name = GetLayoutName(layout1);
                string layout2Name = GetLayoutName(layout2);

                if (!string.IsNullOrEmpty(layout1Name) && !string.IsNullOrEmpty(layout2Name) &&
                    string.Equals(layout1Name, layout2Name, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }

                // Compare by placeholder structure
                return HaveSimilarPlaceholders(layout1, layout2);
            }
            catch
            {
                return false;
            }
        }

        private static bool HaveSimilarPlaceholders(SlideLayoutPart layout1, SlideLayoutPart layout2)
        {
            try
            {
                var placeholders1 = GetPlaceholderTypes(layout1);
                var placeholders2 = GetPlaceholderTypes(layout2);

                return placeholders1.SequenceEqual(placeholders2);
            }
            catch
            {
                return false;
            }
        }

        private static List<PlaceholderValues> GetPlaceholderTypes(SlideLayoutPart layout)
        {
            var placeholders = new List<PlaceholderValues>();

            if (layout?.SlideLayout?.CommonSlideData?.ShapeTree == null)
            {
                return placeholders;
            }

            foreach (var shape in layout.SlideLayout.CommonSlideData.ShapeTree.Elements<Shape>())
            {
                var placeholder = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape;
                if (placeholder?.Type != null)
                {
                    placeholders.Add(placeholder.Type.Value);
                }
            }

            return placeholders.OrderBy(x => x).ToList();
        }

        private static void AddSlideToPresentation(PresentationPart presentationPart, SlidePart slidePart, int? position)
        {
            SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;
            if (slideIdList == null)
            {
                slideIdList = new SlideIdList();
                presentationPart.Presentation.SlideIdList = slideIdList;
            }

            uint newId = GetNextSlideId(slideIdList);

            SlideId newSlideId = new SlideId
            {
                Id = newId,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            };

            if (position.HasValue && position.Value >= 0)
            {
                var slides = slideIdList.Elements<SlideId>().ToList();
                if (position.Value < slides.Count)
                {
                    slideIdList.InsertBefore(newSlideId, slides[position.Value]);
                    return;
                }
            }

            slideIdList.Append(newSlideId);
        }

        private static uint GetNextSlideId(SlideIdList slideIdList)
        {
            if (slideIdList == null || !slideIdList.Any())
                return 256;

            uint maxId = slideIdList.Elements<SlideId>()
                .Where(slideId => slideId.Id != null && slideId.Id.HasValue)
                .Select(slideId => slideId.Id.Value)
                .DefaultIfEmpty((uint)255)
                .Max();

            return Math.Max(256, maxId + 1);
        }

        private static SlidePart GetSlidePartByIndex(PresentationPart presentationPart, int slideIndex)
        {
            if (presentationPart?.Presentation?.SlideIdList == null)
                return null;

            var slideIds = presentationPart.Presentation.SlideIdList.Elements<SlideId>().ToList();

            if (slideIndex < 0 || slideIndex >= slideIds.Count)
                return null;

            SlideId slideId = slideIds[slideIndex];

            if (slideId.RelationshipId != null)
            {
                try
                {
                    return (SlidePart)presentationPart.GetPartById(slideId.RelationshipId);
                }
                catch (ArgumentOutOfRangeException)
                {
                    return null;
                }
            }

            return null;
        }
    }
}
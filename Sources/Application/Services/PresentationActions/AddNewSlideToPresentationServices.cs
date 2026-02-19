using Application.Constants;
using Application.Services.Common;
using Application.Services.PresentationActions.Interfaces;
using Application.Services.PresentationActions.Models;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

using Microsoft.Extensions.Logging;

using Drawing = DocumentFormat.OpenXml.Drawing;

namespace Application.Services.PresentationActions
{
    public class AddNewSlideToPresentationServices(ICopySlideServices copySlideServices) : IAddNewSlideToPresentationServices
    {
        private readonly ICopySlideServices _copySlideServices = copySlideServices;

        public IEnumerable<string> AddNewSlidesByLayout(PresentationDocument presentationDocument, NewSlideData[] slidesData, ILogger logger)
        {
            List<string> result = [];

            foreach (var slideData in slidesData)
            {
                var errors = AddNewSlideByLayout(presentationDocument, slideData, logger);
                if (errors != null)
                {
                    result.AddRange(errors);
                }
            }

            return result;
        }

        public IEnumerable<string>? AddNewSlideByLayout(PresentationDocument presentationDocument, NewSlideData newSlideData, ILogger logger)
        {
            var presentationPart = presentationDocument.PresentationPart;

            // Find SlideLayoutPart
            var slideLayoutPart = PresentationCommonServices.FindSlideLayoutPart(presentationPart, newSlideData.ThemeName, newSlideData.LayoutName);
            if (slideLayoutPart == null)
            {
                return ["Specified layout was not found."];
            }

            // Create a new slide
            var slidePart = AddSlide(presentationPart, slideLayoutPart);

            // Update text for shapes: Title, SubTitle, and Body
            UpdateTextInShapes(slidePart, newSlideData.TitleText, newSlideData.SubTitleText, newSlideData.BodyText);

            // Add Automatically generated label
            //PresentationCommonServices.AddAutomaticallyGeneratedLabel(slidePart);

            // Add comment to the slide
            PresentationCommonServices.AddNoteToSlide(presentationPart, slidePart, newSlideData.CommentMessage);

            // Save changes
            presentationDocument.Save();

            // Validate the document
            var errors = PresentationCommonServices.ValidateDocument(presentationDocument, logger);

            return errors;
        }

        private static SlidePart AddSlide(PresentationPart presentationPart, SlideLayoutPart layoutPart)
        {
            var slidePart = presentationPart.AddNewPart<SlidePart>();

            // Clone shapes from Layout
            slidePart.Slide = new Slide(
                new CommonSlideData(
                    (ShapeTree)layoutPart.SlideLayout.CommonSlideData.ShapeTree.Clone()
                )
            );

            // Copy dependencies
            CopyRelationships(layoutPart, slidePart);

            // Update resource IDs
            UpdateRelationshipIds(slidePart, layoutPart);

            // Attach LayoutPart
            slidePart.AddPart(layoutPart);

            // Add slide to the SlideId list
            AddSlideId(presentationPart, slidePart);

            return slidePart;
        }

        private static void UpdateTextInShapes(SlidePart slidePart, string titleText, string subTitleText, string[] bodyText)
        {
            var shapes = slidePart.Slide.Descendants<Shape>().ToArray();
            //foreach (var shape in shapes)
            for (int index = shapes.Count() -1; index >= 0; index--)
            {
                var shape = shapes[index];

                var placeholder = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape;

            if (placeholder != null && placeholder.Type != null)
                {
                    // Handling Placeholder types: Title, SubTitle, Body
                    if (placeholder.Type == PlaceholderValues.Title || placeholder.Type == PlaceholderValues.CenteredTitle)
                    {
                        if (!string.IsNullOrWhiteSpace(titleText))
                        {
                            UpdateShapeText(shape, titleText);
                        }
                    }
                    else if (placeholder.Type == PlaceholderValues.SubTitle)
                    {
                        if (!string.IsNullOrWhiteSpace(subTitleText))
                        {
                            UpdateShapeText(shape, subTitleText);
                        }
                    }
                    else if (placeholder.Type == PlaceholderValues.Body)
                    {
                        RemoveShapeBodyText(shape);
                        if (!string.IsNullOrWhiteSpace(bodyText[0]))
                        {
                            UpdateShapeBodyText(shape, bodyText);
                        }
                    }
                    else if (placeholder.Type == PlaceholderValues.Picture)
                    {
                        //UpdateShapeImage(slidePart, shape, string.Empty);
                    }
                    else if (placeholder.Type == PlaceholderValues.DateAndTime)
                    {
                        //UpdateShapeText(shape, DateTime.Now.Date.ToString("dd MMMM yyyy"));
                    }
                }
                else
                {
                    if (AddNewSlideServicesConstants.InnerTextsBodyShape.Any(s => shape.InnerText.StartsWith(s)))
                    {
                        UpdateShapeBodyText(shape, bodyText);
                    }
                }
            }
        }

        private static void UpdateShapeImage(SlidePart slidePart, Shape shape, string? imagePath)
        {
            if (string.IsNullOrEmpty(imagePath))
            {
                List<string> images = PresentationCommonServices.GetPicturesList();

                Random rnd = new(DateTime.Now.Microsecond);
                int index = rnd.Next(images.Count);

                imagePath = Path.Combine(AppContext.BaseDirectory, "Resources", images[index]);
            }

            if (string.IsNullOrEmpty(imagePath))
            {
                throw new ArgumentNullException(nameof(imagePath), "Image path cannot be null or empty.");
            }

            if (!File.Exists(imagePath))
            {
                throw new FileNotFoundException($"Image not found at {imagePath}");
            }

            // Add image part to the slide
            ImagePart imagePart;
            string ext = Path.GetExtension(imagePath).ToLowerInvariant();

            imagePart = ext switch
            {
                ".png" => slidePart.AddImagePart(ImagePartType.Png),
                ".jpg" or ".jpeg" => slidePart.AddImagePart(ImagePartType.Jpeg),
                ".gif" => slidePart.AddImagePart(ImagePartType.Gif),
                ".bmp" => slidePart.AddImagePart(ImagePartType.Bmp),
                ".tif" or ".tiff" => slidePart.AddImagePart(ImagePartType.Tiff),
                _ => throw new NotSupportedException($"Unsupported image format: {ext}")
            };

            using (FileStream stream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
            {
                stream.Position = 0;
                imagePart.FeedData(stream);
            }

            string relationshipId = slidePart.GetIdOfPart(imagePart);

            var shapeTree = slidePart.Slide.CommonSlideData.ShapeTree;

            // Remove placeholder marker so PowerPoint stops showing "add picture" icon
            var transform2D = shape.ShapeProperties.Transform2D;
            var nonVisualShapeDrawingProperties = shape.NonVisualShapeProperties.NonVisualDrawingProperties;
            shape.Remove();

            // Create shape with image fill
            Shape imageShape = new(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties() { Id = nonVisualShapeDrawingProperties.Id, Name = nonVisualShapeDrawingProperties.Name },
                    new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties()),
                new ShapeProperties(
                    new Drawing.Transform2D(
                        new Drawing.Offset() { X = transform2D.Offset.X, Y = transform2D.Offset.Y },
                        new Drawing.Extents() { Cx = transform2D.Extents.Cx, Cy = transform2D.Extents.Cy }),
                    new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = Drawing.ShapeTypeValues.Rectangle },
                    new Drawing.BlipFill(
                        new Drawing.Blip() { Embed = relationshipId },
                        new Drawing.Stretch(
                            new Drawing.FillRectangle()))),
                new TextBody(new Drawing.BodyProperties(), new Drawing.ListStyle(), new Drawing.Paragraph())
            );

            shapeTree.AppendChild(imageShape);

            // Save slide
            slidePart.Slide.Save();
        }

        private static void UpdateShapeText(Shape shape, string newText)
        {
            if (shape?.TextBody == null)
            {
                return;
            }

            var paragraphs = shape.TextBody.Elements<Drawing.Paragraph>().ToList();

            // No paragraph: create, append run and exit
            if (paragraphs.Count == 0)
            {
                var para = new Drawing.Paragraph();
                para.Append(new Drawing.Run(new Drawing.Text(newText)));
                shape.TextBody.Append(para);
                return;
            }

            for (int p = 0; p < paragraphs.Count; p++)
            {
                var runs = paragraphs[p].Elements<Drawing.Run>().ToList();
                if (runs.Count > 0)
                {
                    for (int i = 0; i < runs.Count; i++)
                    {
                        runs[i].Text.Text = p == 0 && i == 0 ? newText : string.Empty;
                    }
                }
                else
                {
                    if (p == 0)
                    {
                        paragraphs[p].Append(new Drawing.Run(new Drawing.Text(newText)));
                    }
                }
            }
        }

        private static void RemoveShapeBodyText(Shape shape)
        {
            if (shape.TextBody.Elements<Drawing.Paragraph>().Any())
            {
                int newTextIndex = 0;

                foreach (var paragraph in shape.TextBody.Elements<Drawing.Paragraph>())
                {
                    var runs = paragraph.Descendants<Drawing.Run>();

                    if (runs.Any())
                    {

                        foreach (var run in runs)
                        {
                            run.Text = new Drawing.Text(" ");
                        }
                    }
                }
            }
        }

        private static void UpdateShapeBodyText(Shape shape, string[] newText)
        {
            if (shape?.TextBody == null || newText == null) return;

            var paragraphs = shape.TextBody.Elements<Drawing.Paragraph>().ToList();

            // No paragraph: create, append one run with first line or empty
            if (paragraphs.Count == 0)
            {
                var para = new Drawing.Paragraph();
                string lineText = (newText.Length > 0) ? newText[0] : string.Empty;
                para.Append(new Drawing.Run(new Drawing.Text(lineText)));
                shape.TextBody.Append(para);
                return;
            }

            for (int p = 0; p < paragraphs.Count; p++)
            {
                string lineText = (p == 0 && newText.Length > 0) ? newText[0] : string.Empty;
                var runs = paragraphs[p].Elements<Drawing.Run>().ToList();
                if (runs.Count > 0)
                {
                    for (int i = 0; i < runs.Count; i++)
                    {
                        if (p == 0 && i == 0)
                        {
                            runs[i].Text.Text = lineText;
                        }
                        else
                        {
                            runs[i].Text.Text = string.Empty;
                        }
                    }
                }
                else
                {
                    if (p == 0)
                    {
                        paragraphs[p].Append(new Drawing.Run(new Drawing.Text(lineText)));
                    }
                }
            }
        }

        private static void UpdateRelationshipIds(SlidePart slidePart, SlideLayoutPart layoutPart)
        {
            // Update all relationship IDs in the cloned slide to match their new parts
            foreach (var blip in slidePart.Slide.Descendants<Drawing.Blip>())
            {
                if (blip.Embed == null)
                {
                    continue;
                }

                var oldRelId = blip.Embed.Value;

                // Check if the old relationship ID exists in the layout part
                var oldPart = layoutPart.GetPartById(oldRelId);
                if (oldPart != null)
                {
                    // Copy the part to the new SlidePart and update the relationship ID
                    var newPart = CopyPartToSlidePart(oldPart, slidePart);
                    blip.Embed.Value = slidePart.GetIdOfPart(newPart);
                }
            }
        }

        private static OpenXmlPart CopyPartToSlidePart(OpenXmlPart oldPart, SlidePart slidePart)
        {
            return oldPart switch
            {
                ImagePart imagePart => CopyImagePart(imagePart, slidePart),
                EmbeddedPackagePart embeddedPart => CopyEmbeddedPart(embeddedPart, slidePart),
                _ => throw new NotSupportedException($"Unsupported part type: {oldPart.GetType().Name}")
            };
        }

        private static void AddSlideId(PresentationPart presentationPart, SlidePart slidePart)
        {
            uint maxSlideId = 256;

            var slideIdList = presentationPart.Presentation.SlideIdList;

            if (slideIdList == null)
            {
                slideIdList = new SlideIdList();
                presentationPart.Presentation.SlideIdList = slideIdList;
            }
            else
            {
                maxSlideId = slideIdList.ChildElements.OfType<SlideId>().Max(s => s.Id.Value);
            }

            uint newSlideId = maxSlideId + 1;

            if (newSlideId >= 2147483648)
            {
                List<uint> slideIds = slideIdList.ChildElements.OfType<SlideId>()
                                                .Select(s => s.Id.Value)
                                                .ToList();
                newSlideId = 2147483647;

                while (slideIds.Contains(newSlideId))
                {
                    newSlideId--;
                }
            }

            slideIdList.Append(new SlideId
            {
                Id = newSlideId,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            });
        }

        private static ImagePart CopyImagePart(ImagePart imagePart, SlidePart slidePart)
        {
            var newPart = slidePart.AddImagePart(imagePart.ContentType);
            using var stream = imagePart.GetStream();
            newPart.FeedData(stream);
            return newPart;
        }

        private static EmbeddedPackagePart CopyEmbeddedPart(EmbeddedPackagePart embeddedPart, SlidePart slidePart)
        {
            var newPart = slidePart.AddEmbeddedPackagePart(embeddedPart.ContentType);
            using var stream = embeddedPart.GetStream();
            newPart.FeedData(stream);
            return newPart;
        }

        private static void CopyRelationships(SlideLayoutPart sourcePart, SlidePart targetPart)
        {
            foreach (var part in sourcePart.Parts)
            {
                var oldPart = part.OpenXmlPart;

                // Handle ImagePart (already supported)
                if (oldPart is ImagePart imagePart)
                {
                    CopyImagePart(imagePart, targetPart);
                }
                // Handle EmbeddedPackagePart (already supported)
                else if (oldPart is EmbeddedPackagePart embeddedPart)
                {
                    CopyEmbeddedPart(embeddedPart, targetPart);
                }
                // Handle EmbeddedObjectPart for OLE objects
                else if (oldPart is EmbeddedObjectPart embeddedObjectPart)
                {
                    CopyEmbeddedObjectPart(embeddedObjectPart, targetPart);
                }
                else
                {
                    // Unsupported part types can be logged or ignored
                    Console.WriteLine($"Unsupported part type: {oldPart.GetType().Name}");
                }
            }
        }

        private static EmbeddedObjectPart CopyEmbeddedObjectPart(EmbeddedObjectPart oldPart, SlidePart slidePart)
        {
            // Copy the EmbeddedObjectPart to new SlidePart
            var newPart = slidePart.AddNewPart<EmbeddedObjectPart>(oldPart.ContentType);
            using var stream = oldPart.GetStream();
            newPart.FeedData(stream);
            return newPart;
        }
    }
}

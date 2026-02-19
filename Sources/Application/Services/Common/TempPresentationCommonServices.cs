using Application.Services.PresentationActions.Models;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

using Microsoft.Extensions.Logging;

using Drawing = DocumentFormat.OpenXml.Drawing;

namespace Application.Services.Common
{
    public static class TempPresentationCommonServices
    {
        public static IEnumerable<string> InsertNewSlide(PresentationDocument presentationDocument, NewSlideData newSlideData, int position, ILogger logger)
        {
            NewSlideData slideData;

            PresentationPart? presentationPart = presentationDocument.PresentationPart;

            // 1. Get slide size from presentation
            var slideSize = presentationPart.Presentation.SlideSize;
            long slideWidth = slideSize.Cx;
            long slideHeight = slideSize.Cy;

            // Check if the theme name or layout name is provided. If not, set default values.
            if (string.IsNullOrWhiteSpace(newSlideData.ThemeName) || string.IsNullOrWhiteSpace(newSlideData.LayoutName))
            {
                slideData = newSlideData with { ThemeName = "1_Office Theme", LayoutName = "Office Theme", SlideWidth = slideWidth, SlideHeight = slideHeight };
            }
            else
            {
                slideData = newSlideData with { SlideWidth = slideWidth, SlideHeight = slideHeight };
            }

            SlideLayoutPart? slideLayoutPart = PresentationCommonServices.FindSlideLayoutPart(presentationPart, slideData.ThemeName, slideData.LayoutName);

            // Verify that the presentation document is not empty.
            if (presentationPart is null)
            {
                throw new InvalidOperationException("The presentation document is empty.");
            }

            // Create the slide part and assign the new slide to it.
            SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();

            // Create and initialize a new slide object using the extracted method.
            Slide slide = CreateAndInitializeSlide(slideData, slidePart);

            slidePart.Slide = slide;
            slidePart.Slide.Save();

            // Update the slide ID list in the presentation file.
            SlideIdList? slideIdList = presentationPart.Presentation.SlideIdList;

            if (slideIdList == null)
            {
                slideIdList = new SlideIdList();
                presentationPart.Presentation.SlideIdList = slideIdList;
            }

            uint maxSlideId = 257;
            SlideId? prevSlideId = null;

            OpenXmlElementList slideIds = slideIdList?.ChildElements ?? default;

            foreach (SlideId slideId in slideIds)
            {
                if (slideId.Id != null && slideId.Id > maxSlideId)
                {
                    maxSlideId = slideId.Id;
                }

                position--;
                if (position == 0)
                {
                    prevSlideId = slideId;
                }
            }

            maxSlideId++;

            // Add the slide layout part, if applicable.
            if (slideLayoutPart is not null)
            {
                slidePart.AddPart(slideLayoutPart!);
            }
            else
            {
                // Get the ID of the previous slide.
                SlidePart lastSlidePart;

                if (prevSlideId != null && prevSlideId.RelationshipId != null)
                {
                    lastSlidePart = (SlidePart)presentationPart.GetPartById(prevSlideId.RelationshipId!);
                }
                else
                {
                    string? firstRelId = ((SlideId)slideIds[0])?.RelationshipId;

                    // Throw an exception if the first slide does not contain a relationship ID.
                    if (firstRelId == null)
                    {
                        throw new ArgumentNullException(nameof(firstRelId));
                    }

                    lastSlidePart = (SlidePart)presentationPart.GetPartById(firstRelId);
                }

                // Use the same slide layout from the previous slide.
                if (lastSlidePart.SlideLayoutPart != null)
                {
                    slidePart.AddPart(lastSlidePart.SlideLayoutPart);
                }
            }

            // Insert the new slide into the slide list after the previous slide.
            SlideId newSlideId;
            if (prevSlideId != null)
            {
                newSlideId = slideIdList!.InsertAfter(new SlideId(), prevSlideId);
                newSlideId.Id = maxSlideId;
                newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);
            }
            else
            {
                newSlideId = new SlideId
                {
                    Id = maxSlideId,
                    RelationshipId = presentationPart.GetIdOfPart(slidePart)
                };

                slideIdList!.Append(newSlideId);
            }

            // Add Automatically generated label
            //PresentationCommonServices.AddAutomaticallyGeneratedLabel(slidePart);

            // Add comment to the slide
            PresentationCommonServices.AddNoteToSlide(presentationPart, slidePart, newSlideData.CommentMessage);

            // Validate the presentation and return validation errors, if any.
            return PresentationCommonServices.ValidateDocument(presentationDocument, logger);
        }


        private static GroupShape AddMainShapeGroup(NewSlideData slideData, ref uint drawingObjectId)
        {
            long groupHeight = (long)Math.Round((slideData.SlideHeight * 0.944), MidpointRounding.ToZero);
            long groupWidth = slideData.SlideWidth;
            long shape1Height = (long)Math.Round((groupHeight * 0.55), MidpointRounding.ToZero);
            long shape2Height = groupHeight - shape1Height;

            // Create GroupShape
            GroupShape groupShape = new GroupShape(
                new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties() { Id = drawingObjectId, Name = "MyGroupShape" },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new DocumentFormat.OpenXml.Presentation.GroupShapeProperties(
                    new Drawing.TransformGroup(
                        new Drawing.Offset() { X = 0L, Y = 0L }, // position of group
                        new Drawing.Extents() { Cx = groupWidth, Cy = groupHeight }, // size of group
                        new Drawing.ChildOffset() { X = 0L, Y = 0L },
                        new Drawing.ChildExtents() { Cx = groupWidth, Cy = groupHeight }
                    ))
            );

            drawingObjectId++;

            // Create first shape (violet rectangle)
            Shape violetShape = new Shape(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties() { Id = drawingObjectId, Name = "VioletRect" },
                    new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties()),
                new ShapeProperties(
                    new Drawing.Transform2D(
                        new Drawing.Offset() { X = 0L, Y = 0L },
                        new Drawing.Extents() { Cx = groupWidth, Cy = shape1Height }),
                    new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = Drawing.ShapeTypeValues.Rectangle },
                    new Drawing.SolidFill(new Drawing.RgbColorModelHex() { Val = "5A14C8" })),
                new TextBody(new Drawing.BodyProperties(), new Drawing.ListStyle(), new Drawing.Paragraph())
            );

            drawingObjectId++;

            // Create second shape (blue rectangle)
            Shape blueShape = new Shape(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties() { Id = drawingObjectId, Name = "BlueRect" },
                    new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties()),
                new ShapeProperties(
                    new Drawing.Transform2D(
                        new Drawing.Offset() { X = 0L, Y = shape1Height },
                        new Drawing.Extents() { Cx = groupWidth, Cy = shape2Height }),
                    new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = Drawing.ShapeTypeValues.Rectangle },
                    new Drawing.SolidFill(new Drawing.RgbColorModelHex() { Val = "0030FF" })),
                new TextBody(new Drawing.BodyProperties(), new Drawing.ListStyle(), new Drawing.Paragraph())
            );

            drawingObjectId++;

            // Append shapes to group
            groupShape.Append(violetShape);
            groupShape.Append(blueShape);

            return groupShape;
        }

        private static GroupShape AddFooterShapeGroup(NewSlideData slideData, SlidePart slidePart, ref uint drawingObjectId)
        {
            long groupMainHeight = (long)Math.Round((slideData.SlideHeight * 0.944), MidpointRounding.ToZero);
            long groupHeight = slideData.SlideHeight - groupMainHeight;
            long groupWidth = slideData.SlideWidth;

            // Create GroupShape
            GroupShape groupShape = new GroupShape(
                new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties() { Id = drawingObjectId, Name = "FootGroupShape" },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new GroupShapeProperties(
                    new Drawing.TransformGroup(
                        new Drawing.Offset() { X = 0L, Y = groupMainHeight }, // position of group
                        new Drawing.Extents() { Cx = groupWidth, Cy = groupHeight }, // size of group
                        new Drawing.ChildOffset() { X = 0L, Y = 0L },
                        new Drawing.ChildExtents() { Cx = groupWidth, Cy = groupHeight }
                    ))
            );

            drawingObjectId++;

            //Create first shape(dark blue rectangle)
            Shape darkBlueShape = new Shape(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties() { Id = drawingObjectId, Name = "DarkBlueRect" },
                    new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties()),
                new ShapeProperties(
                    new Drawing.Transform2D(
                        new Drawing.Offset() { X = 0L, Y = 0L },
                        new Drawing.Extents() { Cx = groupWidth, Cy = groupHeight }),
                    new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = Drawing.ShapeTypeValues.Rectangle },
                    new Drawing.SolidFill(new Drawing.RgbColorModelHex() { Val = "001358" })),
                new TextBody(new Drawing.BodyProperties(), new Drawing.ListStyle(), new Drawing.Paragraph())
            );

            drawingObjectId++;

            // Calculate margins
            long marginRight = (long)(groupWidth * 0.02);   // 2% right margin
            long marginVertical = (long)(groupHeight * 0.05); // 5% top and bottom margin

            // Define shape size
            long shapeWidth = (long)(groupWidth * 0.3); // for example, 30% of group width
            long shapeHeight = groupHeight - (2 * marginVertical);

            // Position: aligned to right side with margin
            long offsetX = groupWidth - shapeWidth - marginRight;
            long offsetY = marginVertical;

            // Create text shape
            Shape footerShape = new Shape(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties() { Id = drawingObjectId, Name = "FooterTextRect" },
                    new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Footer })),
                new ShapeProperties(
                    new Drawing.Transform2D(
                        new Drawing.Offset() { X = offsetX, Y = offsetY },
                        new Drawing.Extents() { Cx = shapeWidth, Cy = shapeHeight }),
                    new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = Drawing.ShapeTypeValues.Rectangle }
                ),
                new TextBody(
                    new Drawing.BodyProperties()
                    {
                        Anchor = Drawing.TextAnchoringTypeValues.Center // vertical alignment
                    },
                    new Drawing.ListStyle()
                )
            );

            footerShape?.TextBody?.AppendChild(NewParagraphFull("ADNOC Group", 1400, "en-US", false, "ADNOC Sans", Drawing.TextAlignmentTypeValues.Right));
            footerShape?.TextBody?.AppendChild(NewParagraphFull("Corporate Presentation", 1400, "en-US", false, "ADNOC Sans", Drawing.TextAlignmentTypeValues.Right));

            //    Shape imageShape = AddBackgroundImageShape

            drawingObjectId++;

            offsetX = marginRight;
            offsetY = marginVertical*2;
            shapeHeight = groupHeight - (4 * marginVertical);

            // Append shapes to group
            groupShape.Append(darkBlueShape);
            groupShape.Append(footerShape);
            groupShape.Append(AddImageBackgroundShape(slidePart, "ADNOC.png", offsetX, offsetY, shapeHeight * 2, shapeHeight, ref drawingObjectId));

            return groupShape;
        }

        private static Shape AddImageBackgroundShape(
            SlidePart slidePart,
            string resourceFileName,
            long offsetX,
            long offsetY,
            long width,
            long height,
            ref uint drawingObjectId)
        {
            string imagePath = Path.Combine(AppContext.BaseDirectory, "Resources", resourceFileName);

            if (!File.Exists(imagePath))
                throw new FileNotFoundException($"Image not found at {imagePath}");

            // Add image part
            ImagePart imagePart = slidePart.AddImagePart(ImagePartType.Png);
            using (FileStream stream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
            {
                imagePart.FeedData(stream);
            }

            string relationshipId = slidePart.GetIdOfPart(imagePart);

            // Create shape with image fill
            Shape imageShape = new Shape(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties() { Id = drawingObjectId, Name = "ImageRect" },
                    new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties()),
                new ShapeProperties(
                    new Drawing.Transform2D(
                        new Drawing.Offset() { X = offsetX, Y = offsetY },
                        new Drawing.Extents() { Cx = width, Cy = height }),
                    new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = Drawing.ShapeTypeValues.Rectangle },
                    new Drawing.BlipFill(
                        new Drawing.Blip() { Embed = relationshipId },
                        new Drawing.Stretch(
                            new Drawing.FillRectangle()))),
                new TextBody(new Drawing.BodyProperties(), new Drawing.ListStyle(), new Drawing.Paragraph())
            );

            drawingObjectId++;

            return imageShape;
        }

        private static Shape AddImageBackgroundShape1(
            SlidePart slidePart,
            string resourceFileName,
            long offsetX,
            long offsetY,
            long width,
            long height,
            ref uint drawingObjectId)
        {
            string imagePath = Path.Combine(AppContext.BaseDirectory, "Resources", resourceFileName);

            if (!File.Exists(imagePath))
                throw new FileNotFoundException($"Image not found at {imagePath}");

            // Add image part
            ImagePart imagePart = slidePart.AddImagePart(ImagePartType.Png);
            using (FileStream stream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
            {
                imagePart.FeedData(stream);
            }

            string relationshipId = slidePart.GetIdOfPart(imagePart);


            // Create first shape (violet rectangle)
            Shape imageShape = new Shape(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties() { Id = drawingObjectId, Name = "IcontRect" },
                    new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties()),
                new ShapeProperties(
                    new Drawing.Transform2D(
                        new Drawing.Offset() { X = offsetX, Y = offsetY },
                        new Drawing.Extents() { Cx = width, Cy = height }),
                    new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = Drawing.ShapeTypeValues.Rectangle },
                    new Drawing.SolidFill(new Drawing.RgbColorModelHex() { Val = "FF0000" })),
                new TextBody(new Drawing.BodyProperties(), new Drawing.ListStyle(), new Drawing.Paragraph())
            );

            drawingObjectId++;

            return imageShape;
        }

        // Extracted method for creating and initializing a new slide object.
        public static Slide CreateAndInitializeSlide(NewSlideData slideData, SlidePart slidePart)
        {
            Slide slide = new(new CommonSlideData(new ShapeTree()));
            uint drawingObjectId = 10U;

            // Build the main slide content structure.
            // Specify non-visual properties for the new slide.
            CommonSlideData commonSlideData = slide.CommonSlideData ?? slide.AppendChild(new CommonSlideData());
            ShapeTree shapeTree = commonSlideData.ShapeTree ?? commonSlideData.AppendChild(new ShapeTree());
            NonVisualGroupShapeProperties nonVisualProperties = shapeTree.AppendChild(new NonVisualGroupShapeProperties());
            nonVisualProperties.NonVisualDrawingProperties = new NonVisualDrawingProperties() { Id = 1, Name = "" };
            nonVisualProperties.NonVisualGroupShapeDrawingProperties = new NonVisualGroupShapeDrawingProperties();
            nonVisualProperties.ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties();

            // Specify group shape properties for the slide.
            shapeTree.AppendChild(new GroupShapeProperties());

            shapeTree.Append(AddMainShapeGroup(slideData, ref drawingObjectId));
            shapeTree.Append(AddFooterShapeGroup(slideData, slidePart, ref drawingObjectId));

            long groupHeight = (long)Math.Round((slideData.SlideHeight * 0.944), MidpointRounding.ToZero);
            long groupWidth = slideData.SlideWidth;
            long shape1Height = (long)Math.Round((groupHeight * 0.55), MidpointRounding.ToZero);
            long shape2Height = groupHeight - shape1Height;

            //Add a title to the slide, if provided.
            if (!string.IsNullOrWhiteSpace(slideData.TitleText))
            {

                long marginRight = (long)Math.Round(groupWidth * 0.1, MidpointRounding.ToZero);  // 10% right margin
                long marginVertical = (long)Math.Round(shape1Height * 0.1, MidpointRounding.ToZero); // 10% top and bottom margin

                // Define shape size
                long shapeWidth = (long)Math.Round(groupWidth * 0.8, MidpointRounding.ToZero); // for example, 80% of group width
                long shapeHeight = (long)Math.Round(shape1Height * 0.8, MidpointRounding.ToZero); // for example, 80% of group height

                // Position: aligned to right side with margin
                long offsetX = marginRight;
                long offsetY = marginVertical;

                // Create and configure the title shape.
                Shape titleShape = shapeTree.AppendChild(new Shape());
                drawingObjectId++;

                // Define required properties for the title shape.
                titleShape.NonVisualShapeProperties = new NonVisualShapeProperties(
                    new NonVisualDrawingProperties() { Id = drawingObjectId, Name = "Title" },
                    new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title }));

                titleShape.ShapeProperties = new ShapeProperties(
                    new Drawing.Transform2D(
                        new Drawing.Offset() { X = offsetX, Y = offsetY },
                        new Drawing.Extents() { Cx = shapeWidth, Cy = shapeHeight }));

                // Set the title text.
                titleShape.TextBody = new TextBody(
                    new Drawing.BodyProperties()
                    {
                        Anchor = Drawing.TextAnchoringTypeValues.Center // vertical alignment
                    },
                    new Drawing.ListStyle());
                titleShape.TextBody.AppendChild(NewParagraphFull(slideData.TitleText, 13500, "en-US", false, "ADNOC Sans Light", Drawing.TextAlignmentTypeValues.Center));
            }

            // Add body content to the slide, if provided.
            if (slideData.BodyText != null && slideData.BodyText.Length > 0)
            {
                long marginRight = (long)Math.Round(groupWidth * 0.05, MidpointRounding.ToZero);  // 5% right margin
                long marginVertical = (long)Math.Round(shape1Height * 0.05, MidpointRounding.ToZero); // 5% top and bottom margin

                // Define shape size
                long shapeWidth = (long)Math.Round(groupWidth * 0.9, MidpointRounding.ToZero); // for example, 80% of group width
                long shapeHeight = (long)Math.Round(shape2Height * 0.9, MidpointRounding.ToZero); // for example, 80% of group height

                // Position: aligned to right side with margin
                long offsetX = marginRight;
                long offsetY = shape1Height + marginVertical;

                // Create and configure the body shape.
                Shape bodyShape = shapeTree.AppendChild(new Shape());
                drawingObjectId++;

                // Define required properties for the body shape.
                bodyShape.NonVisualShapeProperties = new NonVisualShapeProperties(new NonVisualDrawingProperties() { Id = drawingObjectId, Name = "Content Placeholder" },
                    new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Index = 1, Type = PlaceholderValues.Body }));

                bodyShape.ShapeProperties = new ShapeProperties(
                    new Drawing.Transform2D(
                        new Drawing.Offset() { X = offsetX, Y = offsetY },
                        new Drawing.Extents() { Cx = shapeWidth, Cy = shapeHeight }));

                // Initialize text body for the shape.
                bodyShape.TextBody = new TextBody(
                    new Drawing.BodyProperties()
                    {
                        Anchor = Drawing.TextAnchoringTypeValues.Center // vertical alignment
                    },
                    new Drawing.ListStyle());

                int textBodyIndex = 0;

                foreach (string paragraphText in slideData.BodyText)
                {
                    bool isBold = false;

                    bodyShape.TextBody.AppendChild(NewParagraphFull(paragraphText, 2400, "en-US", isBold, "ADNOC Sans", Drawing.TextAlignmentTypeValues.Left));
                    //bodyShape.TextBody.AppendChild(NewParagraphFull(paragraphText, 1950, "en-US", isBold, "ADNOC Sans", Drawing.TextAlignmentTypeValues.Left));

                    textBodyIndex++;
                }
            }

            return slide;
        }

        private static Drawing.Paragraph NewParagraphFull(
            string bodyText,
            int fontSize,
            string language,
            bool isBold,
            string typeface,
            Drawing.TextAlignmentTypeValues? align)
        {
            if (align == null)
            {
                align = Drawing.TextAlignmentTypeValues.Left;
            }

            // Paragraph with horizontal alignment
            var paraProps = new Drawing.ParagraphProperties { Alignment = align };
            var paragraph = new Drawing.Paragraph(paraProps);

            // Split text by line breaks
            string[] lines = bodyText.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);

            for (int i = 0; i < lines.Length; i++)
            {
                // Create RunProperties for each line (to avoid shared object issues)
                var runProps = new Drawing.RunProperties();

                // Add the solid fill (color)
                runProps.Append(new Drawing.SolidFill(
                    new Drawing.SchemeColor() { Val = Drawing.SchemeColorValues.Background1 }
                ));

                // Add font properties
                runProps.Append(new Drawing.LatinFont { Typeface = typeface });
                runProps.Append(new Drawing.EastAsianFont { Typeface = typeface });
                runProps.Append(new Drawing.ComplexScriptFont { Typeface = typeface });

                // Set other properties
                runProps.FontSize = fontSize;
                runProps.Language = language;
                runProps.Bold = isBold;

                // Create run with text
                var run = new Drawing.Run(runProps, new Drawing.Text(lines[i]));
                paragraph.Append(run);

                // Add line break after each line except the last one
                if (i < lines.Length - 1)
                {
                    paragraph.Append(new Drawing.Break());
                }
            }

            return paragraph;
        }
    }
}

using System.Text;

using Application.Services.PresentationActions.Models;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2016.Presentation.Command;
using DocumentFormat.OpenXml.Office2021.PowerPoint.Comment;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;

using Microsoft.Extensions.Logging;

using Drawing = DocumentFormat.OpenXml.Drawing;
using Office2021 = DocumentFormat.OpenXml.Office2021.PowerPoint.Comment;

namespace Application.Services.Common
{
    public static class PresentationCommonServices
    {
        public static int SlideIdsCount(PresentationDocument presentationDocument)
        {
            ArgumentNullException.ThrowIfNull(presentationDocument);

            int slideIdsCount = 0;

            PresentationPart? presentationPart = presentationDocument?.PresentationPart;

            if (presentationPart != null)
            {
                Presentation presentation = presentationPart.Presentation;
                SlideIdList? slideIdList = presentation!.SlideIdList;
                slideIdsCount = slideIdList!.Count();
            }

            return slideIdsCount;
        }

        public static IEnumerable<string> ValidateDocument(PresentationDocument presentationDocument, ILogger logger)
        {
#if DEBUG
            var validator = new OpenXmlValidator();
            var errors = validator.Validate(presentationDocument).ToList();

            return new[] { ValidatePresentationStructure(presentationDocument) };
#else
    return Enumerable.Empty<string>();
#endif
        }

        public static string ValidatePresentationStructure(PresentationDocument presentation)
        {
            var validator = new OpenXmlValidator();
            var errors = validator.Validate(presentation).ToList();

            if (!errors.Any())
            {
                return "✅ Presentation passed OpenXml validation.";
            }

            var sb = new StringBuilder();
            sb.AppendLine("⚠️ Validation issues detected:");

            var schemaErrors = errors.Where(e => e.ErrorType == ValidationErrorType.Schema).ToList();
            var semanticErrors = errors.Where(e => e.ErrorType == ValidationErrorType.Semantic).ToList();

            if (schemaErrors.Any())
            {
                sb.AppendLine("\n🔴 Schema Errors:");
                foreach (var err in schemaErrors)
                {
                    sb.AppendLine($"- {err.Description}");
                    sb.AppendLine($"  Part: {err.Part?.Uri}");
                    sb.AppendLine($"  Path: {err.Path}");
                    sb.AppendLine($"  Node: {err.Node?.LocalName}");

                    if (err.Description.Contains("attribute is not declared") && err.Node?.LocalName == "cSld")
                    {
                        sb.AppendLine("  💡 Hint: Remove r:id from <cSld>. Slide should reference layout only via relationships.\n");
                    }
                    else
                    {
                        sb.AppendLine("  💡 Suggestion: Check ID ranges. SlideLayoutId ≥ 257, SlideMasterId ≥ 2147483648.\n");
                    }
                }
            }

            if (semanticErrors.Any())
            {
                sb.AppendLine("\n🟠 Semantic Errors (e.g. missing relationships):");
                foreach (var err in semanticErrors)
                {
                    sb.AppendLine($"- {err.Description}");
                    sb.AppendLine($"  Part: {err.Part?.Uri}");
                    sb.AppendLine($"  Path: {err.Path}");
                    sb.AppendLine($"  Node: {err.Node?.LocalName}");
                    sb.AppendLine($"  💡 Suggestion: Ensure AddPart(...) is called before GetIdOfPart(...), and Save() is called after.\n");
                }
            }

            sb.AppendLine($"🧾 Summary: {schemaErrors.Count} schema errors, {semanticErrors.Count} semantic errors.");
            return sb.ToString();
        }

        public static void AddNoteToSlide(PresentationPart presentationPart, SlidePart slidePart, string noteText)
        {
            //string prefix = "Additional information: ";
            string prefix = "";

            // Get or create NotesSlidePart
            var notesPart = slidePart.NotesSlidePart ?? slidePart.AddNewPart<NotesSlidePart>();
            if (!slidePart.Parts.Any(p => p.OpenXmlPart == notesPart))
            {
                slidePart.AddPart(notesPart);
            }

            // Get or create NotesSlide
            var notesSlide = notesPart.NotesSlide;
            if (notesSlide == null)
            {
                notesSlide = new NotesSlide(
                    new CommonSlideData(new ShapeTree(
                        new NonVisualGroupShapeProperties(
                            new NonVisualDrawingProperties() { Id = 1, Name = "" },
                            new NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        new GroupShapeProperties(),
                        CreateNoteShape($"{prefix} {noteText}")
                    )),
                    new ColorMapOverride(new Drawing.MasterColorMapping())
                );
                notesPart.NotesSlide = notesSlide;
            }
            else
            {
                var shape = notesSlide.Descendants<Shape>()
                    .FirstOrDefault(s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type == PlaceholderValues.Body);

                if (shape == null)
                {
                    notesSlide?.CommonSlideData?.ShapeTree?.Append(CreateNoteShape($"{prefix} {noteText}"));
                }
                else
                {
                    var textBody = shape.TextBody ?? shape.AppendChild(new TextBody(new Drawing.BodyProperties(), new Drawing.ListStyle()));
                    var paragraph = textBody.Elements<Drawing.Paragraph>().FirstOrDefault() ?? textBody.AppendChild(new Drawing.Paragraph());

                    // Remove EndParagraphRunProperties to avoid empty line
                    var endParagraph = paragraph.Descendants<Drawing.EndParagraphRunProperties>().FirstOrDefault();
                    if (endParagraph != null)
                    {
                        paragraph.RemoveChild(endParagraph);
                    }

                    var lastElement = paragraph.ChildElements.LastOrDefault();

                    if (lastElement != null && lastElement.GetType() == typeof(Drawing.Run))
                    {
                        paragraph.Append(new Drawing.Break());
                    }

                    paragraph.Append(new Drawing.Run(new Drawing.Text($"{prefix} {noteText}")));
                    paragraph.Append(new Drawing.EndParagraphRunProperties());
                }
            }

            notesPart.NotesSlide.Save();
            slidePart.Slide.Save();
        }

        private static Shape CreateNoteShape(string noteText)
        {
            return new Shape(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties() { Id = 2, Name = "Notes Placeholder" },
                    new NonVisualShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties(
                        new PlaceholderShape() { Type = PlaceholderValues.Body })),
                new ShapeProperties(),
                new TextBody(
                    new Drawing.BodyProperties(),
                    new Drawing.ListStyle(),
                    new Drawing.Paragraph(
                        new Drawing.Run(new Drawing.Text(noteText)),
                        new Drawing.EndParagraphRunProperties()
                    )
                )
            );
        }

        public static SlideLayoutPart? FindSlideLayoutPart(PresentationPart presentationPart, string themeName, string layoutName)
        {
            return presentationPart.SlideMasterParts
                .Where(master => master.ThemePart?.Theme.Name == themeName)
                .SelectMany(master => master.SlideLayoutParts)
                .FirstOrDefault(layout => layout?.SlideLayout?.CommonSlideData?.Name?.Value == layoutName);
        }

        public static void AddAutomaticallyGeneratedLabel(SlidePart slidePart, long? slideWidth = null)
        {
            long offsetX = 8700000;

            if (slideWidth != null)
            {
                offsetX = (long)(slideWidth*0.75);
            }

            var shapeTree = slidePart?.Slide?.CommonSlideData?.ShapeTree;
            uint shapeId = (uint)(shapeTree!.ChildElements.Count + 1);

            var newShape = new Shape(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties() { Id = shapeId, Name = "RecommendedLabel" },
                    new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Object })
                ),
                new ShapeProperties(
                    new Drawing.Transform2D(
                        new Drawing.Offset() { X = offsetX, Y = 200000 },
                        new Drawing.Extents() { Cx = 2500000, Cy = 300000 }
                    ),
                    new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = Drawing.ShapeTypeValues.Rectangle },
                    // Set dark yellow background
                    new Drawing.SolidFill(
                        new Drawing.RgbColorModelHex() { Val = "FFFF00" } // yellow (hex)
                    )
                ),
                new TextBody(
                    // Set alignment properties for horizontal and vertical text alignment
                    new Drawing.BodyProperties()
                    {
                        Anchor = Drawing.TextAnchoringTypeValues.Center, // Vertical alignment (Center)
                        LeftInset = 0,
                        RightInset = 0,
                        TopInset = 0,
                        BottomInset = 0,
                        Wrap = Drawing.TextWrappingValues.None
                    },
                    new Drawing.ListStyle(),
                    new Drawing.Paragraph(
                        new Drawing.ParagraphProperties()
                        {
                            Alignment = Drawing.TextAlignmentTypeValues.Center // Horizontal alignment (Center)
                        },
                        new Drawing.Run(
                            new Drawing.RunProperties(
                                new Drawing.SolidFill(
                                    new Drawing.RgbColorModelHex() { Val = "000000" } // Text color: Black
                                )
                            )
                            {
                                Language = "en-US",
                                FontSize = 1200,
                                Bold = true
                            },
                            new Drawing.Text("Draft: Automatically generated")
                        ),
                        new Drawing.EndParagraphRunProperties()
                        {
                            Language = "en-US"
                        }
                    )
                )
            );

            shapeTree.Append(newShape);
        }

        public static void ClearAllComments(Presentation? presentation)
        {
            IEnumerable<SlidePart>? slideParts = presentation.PresentationPart?.SlideParts;

            // Remove CommentAuthorPart if exists
            if (presentation.PresentationPart?.CommentAuthorsPart != null)
            {
                presentation.PresentationPart?.DeletePart(presentation.PresentationPart.CommentAuthorsPart);
            }

            if (slideParts != null)
            {
                foreach (SlidePart slidePart in slideParts)
                {
                    // Remove SlideCommentsPart if exists
                    if (slidePart.SlideCommentsPart != null)
                    {
                        slidePart.DeletePart(slidePart.SlideCommentsPart);
                    }

                    // Remove CommentPart if exists
                    if (slidePart.commentParts != null)
                    {
                        foreach (PowerPointCommentPart commentPart in slidePart.commentParts)
                        {
                            slidePart.DeletePart(commentPart);
                        }
                    }
                }
            }

            presentation.Save();
        }

        public static void AddCommentToSlide(PresentationPart presentationPart, SlideId slideId, string authorName, string commentMessage)
        {
            // Ensure the authors part exists
            if (presentationPart.authorsPart == null)
            {
                presentationPart.AddNewPart<PowerPointAuthorsPart>();
            }

            // Ensure the AuthorList exists
            if (presentationPart?.authorsPart?.AuthorList == null)
            {
                presentationPart.authorsPart.AuthorList = new AuthorList();
            }

            // Get or create the author
            AuthorList authors = presentationPart.authorsPart.AuthorList;
            Author author = authors.Elements<Author>().FirstOrDefault(a => a.Name?.Value == authorName);

            if (author == null)
            {
                string authorId = $"{{{Guid.NewGuid()}}}";
                string userId = $"{authorName.Split(' ').FirstOrDefault() ?? "user"}@adnoc.ae::{Guid.NewGuid()}";
                author = new Author()
                {
                    Id = authorId,
                    Name = authorName,
                    Initials = "AG",
                    UserId = userId,
                    ProviderId = string.Empty
                };
                authors.Append(author);
            }

            // Generate a random comment id
            Random ran = new();
            UInt32Value cid = Convert.ToUInt32(ran.Next(100000000, 999999999));

            // Get the relationship id of the slide
            string relId = slideId.RelationshipId;
            SlidePart slidePart = relId != null ? (SlidePart)presentationPart.GetPartById(relId) : presentationPart.SlideParts.First();

            // Get or create the PowerPointCommentPart
            PowerPointCommentPart powerPointCommentPart = slidePart.commentParts.FirstOrDefault() ?? slidePart.AddNewPart<PowerPointCommentPart>();

            // Create the modern comment
            var comment = new DocumentFormat.OpenXml.Office2021.PowerPoint.Comment.Comment(
                new SlideMonikerList(
                    new DocumentMoniker(),
                    new SlideMoniker()
                    {
                        CId = cid,
                        SldId = slideId.Id,
                    }),
                new TextBodyType(
                    new Drawing.BodyProperties(),
                    new Drawing.ListStyle(),
                    new Drawing.Paragraph(
                        new Drawing.Run(
                            new Drawing.RunProperties(
                                new Drawing.SolidFill(
                                    new Drawing.RgbColorModelHex() { Val = "FFFFFF" }
                                )
                            )
                            {
                                Language = "en-US",
                                FontSize = 1200,
                                Bold = true
                            },
                            new Drawing.Text(commentMessage)
                        ),
                        new Drawing.EndParagraphRunProperties() { Language = "en-US" }
                    )))
            {
                Id = $"{{{Guid.NewGuid()}}}",
                AuthorId = author.Id,
                Created = DateTime.Now,
            };

            // Ensure the CommentList exists and add the comment
            powerPointCommentPart.CommentList ??= new DocumentFormat.OpenXml.Office2021.PowerPoint.Comment.CommentList();
            powerPointCommentPart.CommentList.AppendChild(comment);

            // Required URI for modern comments extension
            const string ModernCommentsExtensionUri = "http://schemas.microsoft.com/office/powerpoint/2019/9/slideExtension";

            // Get or create the SlideExtensionList
            SlideExtensionList? presentationExtensionList = slidePart.Slide.ChildElements.OfType<SlideExtensionList>().FirstOrDefault();
            // Create a boolean that determines if this is the slide's first comment

            // If the presentation extension list is null, add one and set this as the first comment for the slide
            if (presentationExtensionList is null)
            {
                slidePart.Slide.AppendChild(new SlideExtensionList());
                presentationExtensionList = slidePart.Slide.ChildElements.OfType<SlideExtensionList>().First();
            }

            // Get or create the SlideExtension with the required URI
            var ext = presentationExtensionList.Elements<SlideExtension>()
                .FirstOrDefault(e => e.Uri?.Value == ModernCommentsExtensionUri);

            if (ext == null)
            {
                ext = new SlideExtension() { Uri = ModernCommentsExtensionUri };
                presentationExtensionList.Append(ext);
            }

            // Remove any existing children and add the CommentRelationship
            ext.RemoveAllChildren();
            ext.Append(new CommentRelationship() { Id = slidePart.GetIdOfPart(powerPointCommentPart) });

            // Save changes to the slide
            slidePart.Slide.Save();
        }

        // Add s comment to the SlidePart
        public static void AddCommentToSlideV2(PresentationPart presentationPart, SlideId slideId, string authorName, string commentMessage)
        {
            if (presentationPart.authorsPart is null)
            {
                presentationPart.AddNewPart<PowerPointAuthorsPart>();
            }

            // Add missing AuthorList if it is null
            if (presentationPart.authorsPart!.AuthorList is null)
            {
                presentationPart.authorsPart.AuthorList = new AuthorList();
            }

            // Get the author or create a new one
            var authors = presentationPart.authorsPart.AuthorList;

            Author? author = presentationPart.authorsPart.AuthorList
                .ChildElements.OfType<Author>().Where(a => a.Name?.Value == authorName).FirstOrDefault();

            if (author is null)
            {
                string authorId = string.Concat("{", Guid.NewGuid(), "}");
                string userId = string.Concat(authorName.Split(" ").FirstOrDefault() ?? "user", "@adnoc.ae::", Guid.NewGuid());
                author = new Author() { Id = authorId, Name = authorName, Initials = "AG", UserId = userId, ProviderId = string.Empty };

                presentationPart.authorsPart.AuthorList.AppendChild(author);
            }

            Random ran = new();
            UInt32Value cid = Convert.ToUInt32(ran.Next(100000000, 999999999));

            // Get the relationship id of the slide if it exists
            string? relId = slideId.RelationshipId;

            // Use the relId to get the slide if it exists, otherwise take the first slide in the sequence
            SlidePart slidePart = relId is not null ? (SlidePart)presentationPart.GetPartById(relId) : presentationPart.SlideParts.First();

            // If the slide part has comments parts take the first PowerPointCommentsPart
            // otherwise add a new one
            PowerPointCommentPart powerPointCommentPart = slidePart.commentParts.FirstOrDefault() ?? slidePart.AddNewPart<PowerPointCommentPart>();

            // Create the comment using the new modern comment class DocumentFormat.OpenXml.Office2021.PowerPoint.Comment.Comment
            var comment = new DocumentFormat.OpenXml.Office2021.PowerPoint.Comment.Comment(
                    new SlideMonikerList(
                        new DocumentMoniker(),
                        new SlideMoniker()
                        {
                            CId = cid,
                            SldId = slideId.Id,
                        }),
                    new TextBodyType(
                        new Drawing.BodyProperties(),
                        new Drawing.ListStyle(),
                        new Drawing.Paragraph(
                            new Drawing.Run(
                                new Drawing.RunProperties(
                                    new Drawing.SolidFill(
                                        new Drawing.RgbColorModelHex() { Val = "FFFFFF" }
                                    )
                                )
                                {
                                    Language = "en-US",
                                    FontSize = 1200,
                                    Bold = true
                                },
                                new Drawing.Text(commentMessage)
                            ),
                            new Drawing.EndParagraphRunProperties() { Language = "en-US" }
                        )))
            {
                Id = string.Concat("{", Guid.NewGuid(), "}"),
                AuthorId = author.Id,
                Created = DateTime.Now,
            };

            // If the comment list does not exist, add one.
            powerPointCommentPart.CommentList ??= new DocumentFormat.OpenXml.Office2021.PowerPoint.Comment.CommentList();
            // Add the comment to the comment list
            powerPointCommentPart.CommentList.AppendChild(comment);

            // Get the presentation extension list if it exists
            SlideExtensionList? presentationExtensionList = slidePart.Slide.ChildElements.OfType<SlideExtensionList>().FirstOrDefault();
            // Create a boolean that determines if this is the slide's first comment
            bool isFirstComment = false;

            // If the presentation extension list is null, add one and set this as the first comment for the slide
            if (presentationExtensionList is null)
            {
                isFirstComment = true;
                slidePart.Slide.AppendChild(new SlideExtensionList());
                presentationExtensionList = slidePart.Slide.ChildElements.OfType<SlideExtensionList>().First();
            }

            // Get the slide extension if it exists
            SlideExtension? presentationExtension = presentationExtensionList.ChildElements.OfType<SlideExtension>().FirstOrDefault();

            // If the slide extension is null, add it and set this as a new comment
            if (presentationExtension is null)
            {
                isFirstComment = true;
                presentationExtensionList.AddChild(new SlideExtension()
                {
                    Uri = Guid.NewGuid().ToString(),

                });
                presentationExtension = presentationExtensionList.ChildElements.OfType<SlideExtension>().First();
            }

            // If this is the first comment for the slide add the comment relationship
            if (isFirstComment)
            {
                presentationExtension.AddChild(new CommentRelationship()
                { Id = slidePart.GetIdOfPart(powerPointCommentPart) });
            }
        }

        private static GroupShape AddMainShapeGroup(NewSlideData slideData)
        {
            long groupHeight = (long)Math.Round((slideData.SlideHeight * 0.944), MidpointRounding.ToZero);
            long groupWidth = slideData.SlideWidth;
            long shape1Height = (long)Math.Round((groupHeight * 0.55), MidpointRounding.ToZero);
            long shape2Height = groupHeight - shape1Height;

            // Create GroupShape
            GroupShape groupShape = new(
                new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties() { Id = 10U, Name = "MyGroupShape" },
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

            // Create first shape (violet rectangle)
            Shape violetShape = new (
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties() { Id = 11U, Name = "VioletRect" },
                    new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties()),
                new ShapeProperties(
                    new Drawing.Transform2D(
                        new Drawing.Offset() { X = 0L, Y = 0L },
                        new Drawing.Extents() { Cx = groupWidth, Cy = shape1Height }),
                    new Drawing.SolidFill(new Drawing.RgbColorModelHex() { Val = "5A14C8" }),
                    new Drawing.Outline(new Drawing.NoFill())),
                new TextBody(new Drawing.BodyProperties(), new Drawing.ListStyle(), new Drawing.Paragraph())
            );

            // Create second shape (blue rectangle)
            Shape blueShape = new (
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties() { Id = 12U, Name = "BlueRect" },
                    new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties()),
                new ShapeProperties(
                    new Drawing.Transform2D(
                        new Drawing.Offset() { X = 0L, Y = shape1Height },
                        new Drawing.Extents() { Cx = groupWidth , Cy = shape2Height }),
                    new Drawing.SolidFill(new Drawing.RgbColorModelHex() { Val = "0030FF" }),
                    new Drawing.Outline(new Drawing.NoFill())),
                new TextBody(new Drawing.BodyProperties(), new Drawing.ListStyle(), new Drawing.Paragraph())
            );

            // 7. Append shapes to group
            groupShape.Append(violetShape);
            groupShape.Append(blueShape);

            return groupShape;
        }

        // Extracted method for creating and initializing a new slide object.
        public static Slide CreateAndInitializeSlide(NewSlideData slideData)
        {
            Slide slide = new (new CommonSlideData(new ShapeTree()));
            uint drawingObjectId = 1;

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

            shapeTree.Append(AddMainShapeGroup(slideData));

            //Add a title to the slide, if provided.
            if (!string.IsNullOrWhiteSpace(slideData.TitleText))
                {
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
                            new Drawing.Offset() { X = 500000, Y = 498476 },
                            new Drawing.Extents() { Cx = 10000000, Cy = 500500 }));

                    // Set the title text.
                    titleShape.TextBody = new TextBody(new Drawing.BodyProperties(), new Drawing.ListStyle());
                    titleShape.TextBody.AppendChild(NewParagraph(slideData.TitleText, 2400, "en-US", true));
                }

            // Add a subtitle to the slide, if provided.
            if (!string.IsNullOrWhiteSpace(slideData.SubTitleText))
            {
                // Create and configure the subtitle shape.
                Shape subTitleShape = shapeTree.AppendChild(new Shape());
                drawingObjectId++;

                // Define required properties for the subtitle shape.
                subTitleShape.NonVisualShapeProperties = new NonVisualShapeProperties(
                    new NonVisualDrawingProperties() { Id = drawingObjectId, Name = "SubTitle" },
                    new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title }));

                subTitleShape.ShapeProperties = new ShapeProperties(
                    new Drawing.Transform2D(
                        new Drawing.Offset() { X = 500000, Y = 950000 },
                        new Drawing.Extents() { Cx = 10000000, Cy = 500500 }));

                // Set the subtitle text.
                subTitleShape.TextBody = new TextBody(new Drawing.BodyProperties(), new Drawing.ListStyle());
                subTitleShape.TextBody.AppendChild(NewParagraph(slideData.SubTitleText, 2400, "en-US", false));
            }

            // Add body content to the slide, if provided.
            if (slideData.BodyText != null && slideData.BodyText.Length > 0)
            {
                // Create and configure the body shape.
                Shape bodyShape = shapeTree.AppendChild(new Shape());
                drawingObjectId++;

                // Define required properties for the body shape.
                bodyShape.NonVisualShapeProperties = new NonVisualShapeProperties(new NonVisualDrawingProperties() { Id = drawingObjectId, Name = "Content Placeholder" },
                    new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Index = 1 }));

                bodyShape.ShapeProperties = new ShapeProperties(
                    new Drawing.Transform2D(
                        new Drawing.Offset() { X = 505000, Y = 1612940 },
                        new Drawing.Extents() { Cx = 10250000, Cy = 2566000 }));

                // Initialize text body for the shape.
                bodyShape.TextBody = new TextBody(new Drawing.BodyProperties(), new Drawing.ListStyle());

                int textBodyIndex = 0;

                foreach (string paragraphText in slideData.BodyText)
                {
                    bool isBold = false;

                    // Make the first paragraph bold if there are multiple lines of text.
                    if (textBodyIndex == 0 && slideData.BodyText.Length > 1)
                    {
                        isBold = true;
                    }

                    bodyShape.TextBody.AppendChild(NewParagraph(paragraphText, 2400, "en-US", isBold));
                    textBodyIndex++;
                }
            }

            return slide;
        }

        // Creates a new paragraph instance with the specified text and properties.
        private static Drawing.Paragraph NewParagraph(string bodyText, int fontSize, string language, bool isBold)
        {
            return new Drawing.Paragraph(
                new Drawing.Run(
                    new Drawing.RunProperties
                    {
                        FontSize = fontSize,
                        Language = language,
                        Bold = isBold
                    },
                    new Drawing.Text() { Text = bodyText }
                )
            );
        }

        public static void AddModernComment(PresentationPart presentationPart, SlideId slideId, string authorName, string commentText)
        {
            SlidePart slidePart = presentationPart.GetPartById(slideId.RelationshipId) as SlidePart ?? throw new Exception("SlidePart not found.");

            // 1. Authors part
            var authorsPart = presentationPart.authorsPart ?? presentationPart.AddNewPart<PowerPointAuthorsPart>();
            authorsPart.AuthorList ??= new AuthorList();

            var author = authorsPart.AuthorList.Elements<Author>().FirstOrDefault(a => a.Name?.Value == authorName);
            if (author == null)
            {
                string authorId = $"{{{Guid.NewGuid()}}}";
                string userId = $"{authorName.Split(' ').FirstOrDefault() ?? "user"}@adnoc.ae::{Guid.NewGuid()}";
                author = new Author()
                {
                    Id = authorId,
                    Name = authorName,
                    Initials = "AG",
                    UserId = userId,
                    ProviderId = string.Empty
                };
                authorsPart.AuthorList.Append(author);
            }
            authorsPart.AuthorList.Save();

            // 2. Modern comment part (SDK will place it in /ppt/comments/)
            var commentPart = presentationPart.GetPartsOfType<PowerPointCommentPart>().FirstOrDefault();
            commentPart ??= presentationPart.AddNewPart<PowerPointCommentPart>();

            // 3. Comment list
            var commentList = commentPart.CommentList ?? new Office2021.CommentList();
            commentPart.CommentList = commentList;

            // 4. Add comment
            string commentId = $"{{{Guid.NewGuid()}}}";
            var comment = new Office2021.Comment(
                new SlideMonikerList(
                    new DocumentMoniker(),
                    new SlideMoniker()
                    {
                        CId = (UInt32Value)(uint)new Random().Next(100000000, 999999999),
                        SldId = slideId.Id,
                    }),
                new TextBodyType(
                    new Drawing.BodyProperties(),
                    new Drawing.ListStyle(),
                    new Drawing.Paragraph(
                        new Drawing.Run(
                            new Drawing.RunProperties(
                                new Drawing.SolidFill(
                                    new Drawing.RgbColorModelHex() { Val = "FFFFFF" }
                                )
                            )
                            {
                                Language = "en-US",
                                FontSize = 1200,
                                Bold = true
                            },
                            new Drawing.Text(commentText)
                        ),
                        new Drawing.EndParagraphRunProperties() { Language = "en-US" }
                    )))
            {
                Id = commentId,
                AuthorId = author.Id,
                Created = DateTime.UtcNow,
            };

            commentList.Append(comment);
            commentList.Save();

            // 5. Relationship from slide to comment part
            string relId = slidePart.CreateRelationshipToPart(commentPart);

            // 6. Slide extension for modern comments
            const string ModernCommentsExtensionUri = "http://schemas.microsoft.com/office/powerpoint/2019/9/slideExtension";

            // 7. Add slide extension for modern comments
            //const string ModernCommentsExtensionUri = "http://schemas.microsoft.com/office/powerpoint/2019/9/slideExtension";

            // Get the presentation extension list if it exists
            SlideExtensionList? presentationExtensionList = slidePart.Slide.ChildElements.OfType<SlideExtensionList>().FirstOrDefault();
            // Create a boolean that determines if this is the slide's first comment

            // If the presentation extension list is null, add one and set this as the first comment for the slide
            if (presentationExtensionList is null)
            {
                slidePart.Slide.AppendChild(new SlideExtensionList());
                presentationExtensionList = slidePart.Slide.ChildElements.OfType<SlideExtensionList>().First();
            }
            //var extList = slidePart.Slide.ExtensionList ?? new SlideExtensionList();
            //slidePart.Slide.ExtensionList = extList;

            var ext = presentationExtensionList.Elements<SlideExtension>()
                .FirstOrDefault(e => e.Uri?.Value == ModernCommentsExtensionUri);

            if (ext == null)
            {
                ext = new SlideExtension() { Uri = ModernCommentsExtensionUri };
                presentationExtensionList.Append(ext);
            }

            ext.RemoveAllChildren();
            ext.Append(new Office2021.CommentRelationship() { Id = relId });

            slidePart.Slide.Save();
        }

        public static List<string> GetPicturesList()
        {
            string folderPath = Path.Combine(AppContext.BaseDirectory, "Resources");
            string[] extensions = { ".jpg", ".jpeg", ".png", ".bmp", ".gif" };

            List<string> imageFiles = Directory
                .GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly)
                .Where(file => extensions.Contains(Path.GetExtension(file).ToLower()))
                .Select(Path.GetFileName)
                .Where(file => file != "ADNOC.png")
                .ToList();

            return imageFiles;
        }
    }
}

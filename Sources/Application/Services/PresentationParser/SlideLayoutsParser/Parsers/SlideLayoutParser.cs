using Application.Services.PresentationParser.Models;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

using Drawing = DocumentFormat.OpenXml.Drawing;

namespace Application.Services.PresentationParser.SlideLayoutsParser.Parsers
{
    /// <summary>
    /// Parses a SlideLayoutPart into a domain model. 
    /// Responsible for extracting text-bearing shapes and computing basic metrics.
    /// </summary>
    internal static class SlideLayoutParser
    {
        public static SlideLayoutInfoDomainModel Parse(
            SlideLayoutPart layoutPart,
            SlideMasterPart masterPart,
            PresentationPart presPart,
            int index)
        {
            // Resolve human-friendly layout name with fallbacks.
            var layoutName = Utils.OpenXmlHelpers.GetLayoutName(layoutPart, index);

            // Heuristic layout id (first shape id in layout), consistent with previous behavior.
            uint? layoutId =
                layoutPart.SlideLayout.CommonSlideData?
                    .ShapeTree?
                    .Elements<Shape>()?
                    .FirstOrDefault()?
                    .NonVisualShapeProperties?
                    .NonVisualDrawingProperties?
                    .Id?.Value;

            // Extract shapes with text
            var shapes = ParseShapes(layoutPart, masterPart, presPart);

            return new SlideLayoutInfoDomainModel(layoutName, layoutId, shapes);
        }

        /// <summary>
        /// Extract and transform shapes (text-bearing) from the layout into domain models.
        /// </summary>
        private static IReadOnlyList<ShapeInfoDomainModel> ParseShapes(
            SlideLayoutPart layoutPart,
            SlideMasterPart masterPart,
            PresentationPart presPart)
        {
            var result = new List<ShapeInfoDomainModel>();

            var shapes = layoutPart.SlideLayout.CommonSlideData?.ShapeTree?.Elements<Shape>();
            if (shapes is null)
            {
                return result;
            }

            foreach (var shape in shapes)
            {
                // Skip shapes without text body
                if (shape.TextBody == null)
                {
                    continue;
                }

                // Quick check for presence of text content (runs or fields)
                bool hasText = shape.TextBody
                    .Elements<Drawing.Paragraph>()
                    .Any(p =>
                        p.ChildElements.OfType<Drawing.Run>().Any() ||
                        p.ChildElements.OfType<Drawing.Field>().Any());

                if (!hasText) continue;

                // Resolve placeholder semantics from layout/master
                var placeholderType = Utils.OpenXmlHelpers.ResolveType(shape, masterPart);
                var placeholderIndex = shape.NonVisualShapeProperties?
                    .ApplicationNonVisualDrawingProperties?
                    .PlaceholderShape?
                    .Index?.Value;

                // Basic identity properties
                string shapeName =
                    shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value
                    ?? "Unnamed Shape";

                uint? shapeId =
                    shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value;

                // Try to map layout placeholder to master placeholder (id)
                var masterShape = Utils.OpenXmlHelpers.FindMasterPlaceholder(masterPart, placeholderType, placeholderIndex);
                var masterShapeId = masterShape?
                    .NonVisualShapeProperties?
                    .NonVisualDrawingProperties?
                    .Id?.Value;

                // Placeholder labels
                var placeholderName = Utils.OpenXmlHelpers.GetPlaceholderName(placeholderType);
                var placeholderRole = Utils.OpenXmlHelpers.GetRole(placeholderType);

                // Inner text area size in points (width/height)
                var (innerWidthPt, innerHeightPt) = Utils.OpenXmlHelpers.GetShapeInnerSizeInPoints(shape);

                // ---- Capacity estimation prerequisites ----
                double? resolvedFontSizePt = null;
                Drawing.ParagraphProperties? resolvedParaProps = null;

                // Find the first text-bearing node and resolve its effective font size.
                foreach (var paragraph in shape.TextBody.Elements<Drawing.Paragraph>())
                {
                    foreach (var node in paragraph.ChildElements)
                    {
                        if (node is Drawing.Run runNode)
                        {
                            var fontSizeInfo = Resolvers.FontAndStyleResolver.ResolveFontAndSize(
                                runNode.RunProperties,
                                paragraph.ParagraphProperties,
                                shape.TextBody,
                                layoutPart,
                                masterPart,
                                presPart,
                                placeholderType,
                                placeholderIndex);

                            resolvedFontSizePt = fontSizeInfo.size;
                            resolvedParaProps = paragraph.ParagraphProperties;
                            break;
                        }
                        else if (node is Drawing.Field)
                        {
                            var fontSizeInfo = Resolvers.FontAndStyleResolver.ResolveFontAndSize(
                                null,
                                paragraph.ParagraphProperties,
                                shape.TextBody,
                                layoutPart,
                                masterPart,
                                presPart,
                                placeholderType,
                                placeholderIndex);

                            resolvedFontSizePt = fontSizeInfo.size;
                            resolvedParaProps = paragraph.ParagraphProperties;
                            break;
                        }
                    }
                    if (resolvedFontSizePt.HasValue) break;
                }

                // ---- Capacity estimation ----
                int? maxLines = null;
                int? maxCharsPerLine = null;

                if (resolvedFontSizePt.HasValue && innerWidthPt.HasValue && innerHeightPt.HasValue)
                {
                    var lineSpacingFactor = Resolvers.FontAndStyleResolver.GetLineSpacingFactor(
                        resolvedParaProps,
                        shape.TextBody,
                        layoutPart,
                        masterPart,
                        placeholderType,
                        resolvedFontSizePt.Value);

                    (maxLines, maxCharsPerLine) =
                        Utils.OpenXmlHelpers.ComputeTextCapacity(
                            innerWidthPt.Value,
                            innerHeightPt.Value,
                            resolvedFontSizePt.Value,
                            lineSpacingFactor,
                            Utils.ParsingDefaults.DefaultAvgCharWidthEm);
                }

                // ---- Extract paragraphs & runs into domain models ----
                var paragraphs = new List<ParagraphInfoDomainModel>();
                int paragraphIndex = 0;

                foreach (var paragraph in shape.TextBody.Elements<Drawing.Paragraph>())
                {
                    var runs = new List<RunInfoDomainModel>();
                    int runIndex = 0;

                    foreach (var node in paragraph.ChildElements)
                    {
                        if (node is Drawing.Field fld)
                        {
                            var fieldText = string.Concat(
                                fld.Descendants<Drawing.Text>().Select(t => t.Text));

                            var fontSizeInfo = Resolvers.FontAndStyleResolver.ResolveFontAndSize(
                                null,
                                paragraph.ParagraphProperties,
                                shape.TextBody,
                                layoutPart,
                                masterPart,
                                presPart,
                                placeholderType,
                                placeholderIndex);

                            runs.Add(new RunInfoDomainModel(
                                runIndex,
                                fieldText,
                                fontSizeInfo.font,
                                fontSizeInfo.size,
                                IsField: true,
                                fld.Type?.Value,
                                fld.Id?.Value));

                            runIndex++;
                        }
                        else if (node is Drawing.Run run)
                        {
                            var fontSizeInfo = Resolvers.FontAndStyleResolver.ResolveFontAndSize(
                                run.RunProperties,
                                paragraph.ParagraphProperties,
                                shape.TextBody,
                                layoutPart,
                                masterPart,
                                presPart,
                                placeholderType,
                                placeholderIndex);

                            runs.Add(new RunInfoDomainModel(
                                runIndex,
                                run.Text?.Text ?? string.Empty,
                                fontSizeInfo.font,
                                fontSizeInfo.size));

                            runIndex++;
                        }
                    }

                    if (runs.Count > 0)
                    {
                        paragraphs.Add(new ParagraphInfoDomainModel(paragraphIndex, runs));
                    }
                    paragraphIndex++;
                }

                // Final shape projection
                result.Add(new ShapeInfoDomainModel(
                    shapeName,
                    shapeId,
                    placeholderIndex,
                    masterShapeId,
                    placeholderName,
                    placeholderType,
                    placeholderRole,
                    maxLines,
                    maxCharsPerLine,
                    paragraphs));
            }

            return result;
        }
    }
}

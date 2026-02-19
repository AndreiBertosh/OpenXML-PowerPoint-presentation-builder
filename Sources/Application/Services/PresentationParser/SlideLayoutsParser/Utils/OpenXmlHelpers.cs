using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

using Drawing = DocumentFormat.OpenXml.Drawing;

namespace Application.Services.PresentationParser.SlideLayoutsParser.Utils
{
    /// <summary>
    /// Shared constants used across parsers/resolvers.
    /// </summary>
    internal static class ParsingDefaults
    {
        // Default fallback values used by the capacity estimator and layout geometry.
        public const double DefaultLineSpacingFactor = 1.2;
        public const double DefaultAvgCharWidthEm = 0.5; // heuristic avg character width in em
        public const long DefaultInsetEmu = 91440;       // ~0.1 inch
        public const double MinCharWidthFactor = 0.3;
        public const double MaxCharWidthFactor = 1.0;
    }

    /// <summary>
    /// Helper methods for placeholder semantics, naming, geometry, theme aliasing, and text capacity.
    /// </summary>
    internal static class OpenXmlHelpers
    {
        // -------------------- Placeholder helpers --------------------

        /// <summary>
        /// True if placeholder type is any title (Title or CenteredTitle).
        /// </summary>
        public static bool IsAnyTitle(PlaceholderValues? t) =>
            t == PlaceholderValues.Title || t == PlaceholderValues.CenteredTitle;

        /// <summary>
        /// Human-readable placeholder type name or "None" if not a placeholder.
        /// </summary>
        public static string GetPlaceholderName(PlaceholderValues? type) =>
            type?.ToString() ?? "None";

        /// <summary>
        /// High-level role name derived from placeholder type (Title, Subtitle, Body, Date, Footer, SlideNumber, Other).
        /// </summary>
        public static string GetRole(PlaceholderValues? type)
        {
            if (type == null) return "Text";
            var name = type.Value.ToString();
            if (name == "Title" || name == "CenteredTitle") return "Title";
            if (name == "SubTitle") return "Subtitle";
            if (name == "Body") return "Body";
            if (name == "Date" || name == "DateAndTime") return "Date";
            if (name == "Footer") return "Footer";
            if (name == "SlideNumber") return "SlideNumber";
            return "Other";
        }

        /// <summary>
        /// Try to locate the corresponding master placeholder shape by index or type.
        /// </summary>
        public static Shape? FindMasterPlaceholder(
            SlideMasterPart masterPart,
            PlaceholderValues? type,
            uint? idx)
        {
            var shapes = masterPart.SlideMaster.CommonSlideData?.ShapeTree?.Elements<Shape>();
            if (shapes == null) return null;

            // Prefer explicit placeholder index mapping if available
            if (idx.HasValue)
            {
                var byIdx = shapes.FirstOrDefault(s =>
                    s.NonVisualShapeProperties?
                        .ApplicationNonVisualDrawingProperties?
                        .PlaceholderShape?
                        .Index?.Value == idx.Value);

                if (byIdx != null) return byIdx;
            }

            // Fall back to type matching (with special handling for any title)
            return type != null
                ? shapes.FirstOrDefault(s =>
                {
                    var t = s.NonVisualShapeProperties?
                        .ApplicationNonVisualDrawingProperties?
                        .PlaceholderShape?
                        .Type?.Value;

                    return IsAnyTitle(type)
                        ? (t == PlaceholderValues.Title || t == PlaceholderValues.CenteredTitle)
                        : (t == type);
                })
                : null;
        }

        /// <summary>
        /// Resolve placeholder type for a layout shape using: explicit layout type -> index mapping to master -> name heuristic.
        /// </summary>
        public static PlaceholderValues? ResolveType(Shape layoutShape, SlideMasterPart masterPart)
        {
            var ph = layoutShape.NonVisualShapeProperties?
                .ApplicationNonVisualDrawingProperties?
                .PlaceholderShape;

            // 1) Direct type on layout shape
            var type = ph?.Type?.Value;
            if (type != null) return type;

            // 2) Try by index: map to master shape and inherit its type
            var idx = ph?.Index?.Value;
            if (idx.HasValue)
            {
                var masterShapes = masterPart.SlideMaster.CommonSlideData?.ShapeTree?.Elements<Shape>();
                var target = masterShapes?.FirstOrDefault(s =>
                {
                    var mph = s.NonVisualShapeProperties?
                        .ApplicationNonVisualDrawingProperties?
                        .PlaceholderShape;
                    return mph?.Index?.Value == idx.Value;
                });

                var masterType = target?
                    .NonVisualShapeProperties?
                    .ApplicationNonVisualDrawingProperties?
                    .PlaceholderShape?
                    .Type?.Value;

                if (masterType != null) return masterType;
            }

            // 3) Heuristic by shape name (rare fallback for shapes without placeholder info)
            var name = layoutShape.NonVisualShapeProperties?
                .NonVisualDrawingProperties?
                .Name?.Value ?? string.Empty;

            if (name.Contains("Title", StringComparison.OrdinalIgnoreCase))
                return PlaceholderValues.Title;

            if (name.Contains("SubTitle", StringComparison.OrdinalIgnoreCase) ||
                name.Contains("Subtitle", StringComparison.OrdinalIgnoreCase))
                return PlaceholderValues.SubTitle;

            if (name.Contains("Body", StringComparison.OrdinalIgnoreCase) ||
                name.Contains("Content", StringComparison.OrdinalIgnoreCase) ||
                name.Contains("Text", StringComparison.OrdinalIgnoreCase))
                return PlaceholderValues.Body;

            // Unknown -> null (treat as normal text box in role helper)
            return null;
        }

        // -------------------- Theme alias & naming helpers --------------------

        /// <summary>
        /// Resolve Office "theme alias" fonts like "+mn-lt" or "+mj-lt" to actual typefaces using ThemePart.
        /// If the typeface is not an alias, returns it as-is.
        /// </summary>
        public static string? ResolveThemeAlias(string? typeface, PresentationPart presPart)
        {
            if (string.IsNullOrEmpty(typeface) || !typeface.StartsWith("+"))
                return typeface;

            var scheme = presPart.ThemePart?.Theme?.ThemeElements?.FontScheme;
            return typeface switch
            {
                "+mj-lt" => scheme?.MajorFont?.GetFirstChild<Drawing.LatinFont>()?.Typeface,
                "+mn-lt" => scheme?.MinorFont?.GetFirstChild<Drawing.LatinFont>()?.Typeface,
                _ => typeface,
            };
        }

        /// <summary>
        /// Return slide master name strictly from Theme.Name (fallback to "Slide Master" if missing).
        /// </summary>
        public static string GetMasterNameFromTheme(SlideMasterPart masterPart)
        {
            var themeName = masterPart?.ThemePart?.Theme?.Name?.Value;
            return string.IsNullOrWhiteSpace(themeName) ? "Slide Master" : themeName!;
            // NOTE: per product decision we do not fall back to cSld/name or URI here.
        }

        /// <summary>
        /// Resolve layout name from (cSld/name) or (type friendly) or (URI) or ordinal fallback.
        /// </summary>
        public static string GetLayoutName(SlideLayoutPart layoutPart, int index)
        {
            // 1) p:cSld/p:name — primary source
            var name = layoutPart.SlideLayout?.CommonSlideData?.Name?.Value;
            if (!string.IsNullOrWhiteSpace(name)) return name!;

            // 2) p:sldLayout/@type -> friendly string (e.g., TitleAndContent -> "Title and Content")
            var type = layoutPart.SlideLayout?.Type?.Value;
            if (type != null)
            {
                return ToFriendlyLayoutName(type.Value);
            }

            // 3) From part URI (e.g., slideLayout3.xml -> slideLayout3)
            var uri = layoutPart.Uri;
            if (uri != null)
            {
                string file = System.IO.Path.GetFileNameWithoutExtension(uri.ToString());
                if (!string.IsNullOrWhiteSpace(file)) return file;
            }

            // 4) Ordinal fallback
            return $"Layout #{index}";
        }

        /// <summary>
        /// Convert SlideLayoutValues enum name (PascalCase) to spaced words.
        /// </summary>
        public static string ToFriendlyLayoutName(SlideLayoutValues layoutType)
        {
            var raw = layoutType.ToString();
            if (string.IsNullOrWhiteSpace(raw)) return "Layout";
            var sb = new System.Text.StringBuilder(raw.Length + 8);
            for (int i = 0; i < raw.Length; i++)
            {
                char c = raw[i];
                if (i > 0 && char.IsUpper(c) && char.IsLower(raw[i - 1])) sb.Append(' ');
                sb.Append(c);
            }
            return sb.ToString();
        }

        // -------------------- Geometry & text capacity --------------------

        /// <summary>
        /// Compute inner text area size (width, height) in points considering text insets (in EMU).
        /// </summary>
        public static (double? innerWidthPt, double? innerHeightPt) GetShapeInnerSizeInPoints(Shape shape)
        {
            var extents = shape.ShapeProperties?.Transform2D?.Extents;
            var cx = extents?.Cx?.Value;
            var cy = extents?.Cy?.Value;
            if (cx == null || cy == null) return (null, null);

            var widthPt = EmuToPoints(cx.Value);
            var heightPt = EmuToPoints(cy.Value);

            var bodyPr = shape.TextBody?.BodyProperties;
            long lIns = bodyPr?.LeftInset?.Value ?? ParsingDefaults.DefaultInsetEmu;
            long rIns = bodyPr?.RightInset?.Value ?? ParsingDefaults.DefaultInsetEmu;
            long tIns = bodyPr?.TopInset?.Value ?? ParsingDefaults.DefaultInsetEmu;
            long bIns = bodyPr?.BottomInset?.Value ?? ParsingDefaults.DefaultInsetEmu;

            return (
                Math.Max(0, widthPt - EmuToPoints(lIns + rIns)),
                Math.Max(0, heightPt - EmuToPoints(tIns + bIns))
            );
        }

        /// <summary>
        /// Convert EMU (English Metric Units) to points.
        /// </summary>
        public static double EmuToPoints(long emu) => emu / 12700.0;

        /// <summary>
        /// Estimate maximum number of lines and characters per line for a given text area and font metrics.
        /// </summary>
        public static (int maxLines, int maxCharsPerLine) ComputeTextCapacity(
            double innerWidthPt,
            double innerHeightPt,
            double fontSizePt,
            double lineSpacingFactor,
            double avgCharWidthEm)
        {
            // Compute line height including spacing; at least the font size.
            var lineHeightWithSpacing =
                Math.Max(fontSizePt, fontSizePt * Math.Max(0.5, lineSpacingFactor));

            // If the area is smaller than a single line -> capacity is zero
            if (innerHeightPt < fontSizePt) return (0, 0);

            // Max lines: first line fits fontSizePt; remaining space divided by line height + spacing.
            var maxLines =
                1 + (int)Math.Floor((innerHeightPt - fontSizePt) / lineHeightWithSpacing);

            // Average character width in points (bounded heuristics)
            var avgCharWidthPt = fontSizePt * Math.Max(
                ParsingDefaults.MinCharWidthFactor,
                Math.Min(avgCharWidthEm, ParsingDefaults.MaxCharWidthFactor));

            var maxCharsPerLine = (int)Math.Floor(innerWidthPt / avgCharWidthPt);

            return (Math.Max(0, maxLines), Math.Max(0, maxCharsPerLine));
        }
    }
}

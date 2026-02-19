using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

using Drawing = DocumentFormat.OpenXml.Drawing;

namespace Application.Services.PresentationParser.SlideLayoutsParser.Resolvers
{
    /// <summary>
    /// Resolves effective font and size using a well-defined fallback chain:
    /// Run -> Paragraph -> TextBody -> Placeholder (master) -> Layout/Master text styles -> Theme alias.
    /// Also provides methods to resolve line spacing factor.
    /// </summary>
    internal static class FontAndStyleResolver
    {
        /// <summary>
        /// Resolve (font, size) from run/paragraph/body/layout/master/theme chain.
        /// Returns nulls if nothing was found (should be handled by caller).
        /// </summary>
        public static (string? font, double? size) ResolveFontAndSize(
            Drawing.RunProperties? runProps,
            Drawing.ParagraphProperties? paraProps,
            TextBody textBody,
            SlideLayoutPart layoutPart,
            SlideMasterPart masterPart,
            PresentationPart presPart,
            PlaceholderValues? placeholderType,
            uint? placeholderIndex)
        {
            string? font = null;
            double? size = null;

            // 1) Run-level (strongest)
            font ??= runProps?.GetFirstChild<Drawing.LatinFont>()?.Typeface;
            size ??= runProps?.FontSize?.Value / 100.0;

            // 2) Paragraph default run properties
            if (font == null || size == null)
            {
                var def = paraProps?.GetFirstChild<Drawing.DefaultRunProperties>();
                font ??= def?.GetFirstChild<Drawing.LatinFont>()?.Typeface;
                size ??= def?.FontSize?.Value / 100.0;
            }

            // 3) TextBody default run (closest container)
            if (font == null || size == null)
            {
                var fromBody = GetFromTextBody(textBody);
                font ??= fromBody.font;
                size ??= fromBody.size;
            }

            // 4) Placeholder (from master shape text body)
            if (font == null)
            {
                font = GetFontFromMasterPlaceholder(masterPart, placeholderType, placeholderIndex, presPart);
            }

            // 5) Title size from text styles (layout's master)
            if (size == null && Utils.OpenXmlHelpers.IsAnyTitle(placeholderType))
            {
                size = GetTitleSizeFromTextStyles(layoutPart);
            }

            // 6) Master text styles (body style) + theme alias
            if (font == null || size == null)
            {
                var fromStyles = GetFromMasterTextStyles(masterPart.SlideMaster.TextStyles, placeholderType, presPart);
                font ??= fromStyles.font;
                size ??= fromStyles.size;
            }

            return (font, size);
        }

        /// <summary>
        /// Try to read font+size from TextBody default run properties.
        /// </summary>
        private static (string? font, double? size) GetFromTextBody(TextBody body)
        {
            // Note: Descendants is recursive; first entry with both a LatinFont and FontSize is usually the most relevant.
            var def = body.Descendants<Drawing.DefaultRunProperties>()
                .FirstOrDefault(r => r.FontSize != null || r.GetFirstChild<Drawing.LatinFont>() != null);

            return def == null
                ? (null, null)
                : (def.GetFirstChild<Drawing.LatinFont>()?.Typeface, def.FontSize?.Value / 100.0);
        }

        /// <summary>
        /// Return font typeface from the master placeholder's default run properties, resolving theme aliases.
        /// </summary>
        private static string? GetFontFromMasterPlaceholder(
            SlideMasterPart masterPart,
            PlaceholderValues? type,
            uint? idx,
            PresentationPart presPart)
        {
            var target = Utils.OpenXmlHelpers.FindMasterPlaceholder(masterPart, type, idx);
            var def = target?.TextBody?.Descendants<Drawing.DefaultRunProperties>().FirstOrDefault();
            var font = def?.GetFirstChild<Drawing.LatinFont>()?.Typeface;
            return Utils.OpenXmlHelpers.ResolveThemeAlias(font, presPart);
        }

        /// <summary>
        /// Return title font size from layout's master text styles (TitleStyle).
        /// </summary>
        private static double? GetTitleSizeFromTextStyles(SlideLayoutPart layoutPart)
        {
            var styles = layoutPart.SlideMasterPart?.SlideMaster?.TextStyles;
            var def = styles?.TitleStyle?.Descendants<Drawing.Level1ParagraphProperties>()
                .Select(p => p.GetFirstChild<Drawing.DefaultRunProperties>())
                .FirstOrDefault(r => r?.FontSize != null);

            return def?.FontSize?.Value / 100.0;
        }

        /// <summary>
        /// Get (font,size) from master text styles (body style), resolving theme aliases.
        /// Skip for titles (they use TitleStyle instead).
        /// </summary>
        private static (string? font, double? size) GetFromMasterTextStyles(
            TextStyles? styles,
            PlaceholderValues? type,
            PresentationPart presPart)
        {
            if (styles == null || Utils.OpenXmlHelpers.IsAnyTitle(type))
            {
                return (null, null);
            }

            var def = styles.BodyStyle?.Descendants<Drawing.Level1ParagraphProperties>()
                .Select(p => p.GetFirstChild<Drawing.DefaultRunProperties>())
                .FirstOrDefault();

            if (def == null)
            {
                return (null, null);
            }

            var font = Utils.OpenXmlHelpers.ResolveThemeAlias(
                def.GetFirstChild<Drawing.LatinFont>()?.Typeface, presPart);
            var size = def.FontSize?.Value / 100.0;

            return (font, size);
        }

        /// <summary>
        /// Resolve line spacing factor using Paragraph -> TextBody ListStyle -> Master TextStyles fallbacks.
        /// </summary>
        public static double GetLineSpacingFactor(
            Drawing.ParagraphProperties? paraProps,
            TextBody textBody,
            SlideLayoutPart layoutPart,
            SlideMasterPart masterPart,
            PlaceholderValues? placeholderType,
            double fontSizePt)
        {
            // 1) Paragraph-level override
            var fromPara = TryGetFactorFromLineSpacing(paraProps?.LineSpacing, fontSizePt);
            if (fromPara.HasValue)
            {
                return fromPara.Value;
            }

            // 2) TextBody ListStyle (first applicable level)
            var fromList = TryGetFactorFromListStyle(textBody.ListStyle, fontSizePt);
            if (fromList.HasValue)
            {
                return fromList.Value;
            }

            // 3) Master text styles (TitleStyle/BodyStyle depending on placeholder type)
            var fromMaster = TryGetFactorFromTextStyles(masterPart.SlideMaster.TextStyles, placeholderType, fontSizePt);
            return fromMaster ?? Utils.ParsingDefaults.DefaultLineSpacingFactor;
        }

        /// <summary>
        /// Convert OpenXML line spacing specification to a relative factor (>= 0.5).
        /// Supports percent or absolute points.
        /// </summary>
        private static double? TryGetFactorFromLineSpacing(Drawing.LineSpacing? lnSpc, double fontSizePt)
        {
            if (lnSpc == null || fontSizePt <= 0)
            {
                return null;
            }

            var pctVal = lnSpc.SpacingPercent?.Val?.Value;
            if (pctVal.HasValue)
            {
                return Math.Max(0.5, pctVal.Value / 100000.0);
            }

            var ptsVal = lnSpc.SpacingPoints?.Val?.Value;
            if (ptsVal.HasValue)
            {
                return Math.Max(0.5, (ptsVal.Value / 100.0) / fontSizePt);
            }

            return null;
        }

        /// <summary>
        /// Inspect ListStyle levels and return the first available line spacing factor.
        /// Avoids 'dynamic': checks Level1..Level9 paragraph properties explicitly.
        /// </summary>
        private static double? TryGetFactorFromListStyle(Drawing.ListStyle? listStyle, double fontSizePt)
        {
            if (listStyle == null || fontSizePt <= 0)
            {
                return null;
            }

            // List of paragraph property levels supported by OpenXML
            Type[] levels =
            [
                typeof(Drawing.Level1ParagraphProperties),
                typeof(Drawing.Level2ParagraphProperties),
                typeof(Drawing.Level3ParagraphProperties),
                typeof(Drawing.Level4ParagraphProperties),
                typeof(Drawing.Level5ParagraphProperties),
                typeof(Drawing.Level6ParagraphProperties),
                typeof(Drawing.Level7ParagraphProperties),
                typeof(Drawing.Level8ParagraphProperties),
                typeof(Drawing.Level9ParagraphProperties)
            ];

            foreach (var levelType in levels)
            {
                var props = listStyle.Descendants()
                    .FirstOrDefault(p => p.GetType() == levelType);

                if (props == null)
                {
                    continue;
                }

                // dynamic is required because each LevelN class has its own LineSpacing property
                dynamic dyn = props;

                Drawing.LineSpacing? spacing = dyn.LineSpacing as Drawing.LineSpacing;
                var factor = TryGetFactorFromLineSpacing(spacing, fontSizePt);

                if (factor.HasValue)
                {
                    return factor.Value;
                }
            }

            return null;
        }

        /// <summary>
        /// Read line spacing factor from master text styles, switching between TitleStyle and BodyStyle.
        /// </summary>
        private static double? TryGetFactorFromTextStyles(TextStyles? styles, PlaceholderValues? type, double fontSizePt)
        {
            if (styles == null || fontSizePt <= 0)
            {
                return null;
            }

            var lvl = Utils.OpenXmlHelpers.IsAnyTitle(type)
                ? styles.TitleStyle?.Descendants<Drawing.Level1ParagraphProperties>().FirstOrDefault()
                : styles.BodyStyle?.Descendants<Drawing.Level1ParagraphProperties>().FirstOrDefault();

            return TryGetFactorFromLineSpacing(lvl?.LineSpacing, fontSizePt);
        }
    }
}

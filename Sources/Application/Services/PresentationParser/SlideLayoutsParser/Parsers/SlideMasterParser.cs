using Application.Services.PresentationParser.Models;

using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;

namespace Application.Services.PresentationParser.SlideLayoutsParser.Parsers
{

    /// <summary>
    /// Parses a SlideMasterPart into a domain model, delegating layout/shape work.
    /// </summary>
    internal static class SlideMasterParser
    {
        /// <summary>
        /// Parse a single SlideMasterPart and produce SlideMasterDomainModel.
        /// </summary>
        public static SlideMasterInfoDomainModel Parse(SlideMasterPart masterPart, PresentationPart presPart)
        {
            // Master name: only from Theme.Name (per current product decision)
            var masterName = Utils.OpenXmlHelpers.GetMasterNameFromTheme(masterPart);

            // A heuristic masterId (first shape id in master), matching previous behavior.
            uint? masterId =
                masterPart.SlideMaster.CommonSlideData?
                    .ShapeTree?
                    .Elements<Shape>()?
                    .FirstOrDefault()?
                    .NonVisualShapeProperties?
                    .NonVisualDrawingProperties?
                    .Id?.Value;

            // Parse layouts
            var layouts = new List<SlideLayoutInfoDomainModel>();
            foreach (var layoutPart in masterPart.SlideLayoutParts)
            {
                var layoutModel = SlideLayoutParser.Parse(layoutPart, masterPart, presPart, index: layouts.Count + 1);
                layouts.Add(layoutModel);
            }

            return new SlideMasterInfoDomainModel(masterName, masterId, layouts);
        }
    }
}

using Application.Services.PresentationActions.Interfaces;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace Application.Services.PresentationActions
{
    public class SlideMasterCopyService : ISlideMasterCopyService
    {
        /// <summary>
        /// Core routine: copies a Slide Master (with all layouts and dependencies)
        /// from source PresentationPart to target PresentationPart by Theme.Name.
        /// Returns the relationship id of the copied or existing master in the target.
        /// </summary>
        public Task<string> CopySlideMasterByThemeName(
            PresentationPart sourcePresPart,
            PresentationPart targetPresPart,
            string themeName,
            bool skipIfExistsInTarget = true)
        {
            ArgumentNullException.ThrowIfNull(sourcePresPart);

            ArgumentNullException.ThrowIfNull(targetPresPart);

            if (string.IsNullOrWhiteSpace(themeName))
            {
                throw new ArgumentException("Theme name is required.", nameof(themeName));
            }

            // Ensure root Presentation and SlideMasterIdList exist on the target
            EnsurePresentationAndMasterIdList(targetPresPart);

            // 1) If target already has a master with the same theme, optionally skip
            var existingTargetMaster = TryFindSlideMasterByTheme(targetPresPart, themeName);
            if (existingTargetMaster != null && skipIfExistsInTarget)
            {
                return Task.FromResult(targetPresPart.GetIdOfPart(existingTargetMaster));
            }

            // 2) Find source SlideMaster by Theme.Name
            var sourceMaster = TryFindSlideMasterByTheme(sourcePresPart, themeName);
            if (sourceMaster == null)
            {
                throw new InvalidOperationException(
                    $"Slide Master with Theme.Name='{themeName}' was not found in the source presentation.");
            }

            // 3) Clone the source master into the target (brings layouts, theme, media, charts, etc.)
            var clonedMasterPart = targetPresPart.AddPart(sourceMaster);

            // 4) Register the cloned master in the target's SlideMasterIdList
            var relId = targetPresPart.GetIdOfPart(clonedMasterPart);
            var smList = targetPresPart.Presentation.SlideMasterIdList;

            uint newId = 1;
            if (smList.HasChildren)
            {
                var maxId = smList.Elements<SlideMasterId>()
                    .Select(e => (uint)e.Id)
                    .DefaultIfEmpty(0u)
                    .Max();

                newId = checked(maxId + 1);
            }

            var newSlideMasterId = new SlideMasterId
            {
                Id = new UInt32Value(newId),
                RelationshipId = relId
            };

            smList.Append(newSlideMasterId);

            // Persist the change
            targetPresPart.Presentation.Save();

            return Task.FromResult(relId);
        }

        /// <summary>
        /// Helper: find a SlideMasterPart in a PresentationPart by Theme.Name.
        /// </summary>
        private static SlideMasterPart TryFindSlideMasterByTheme(PresentationPart presPart, string themeName)
        {
            return presPart
                .SlideMasterParts
                .FirstOrDefault(mp => string.Equals(
                    mp.ThemePart?.Theme?.Name?.Value,
                    themeName,
                    StringComparison.Ordinal));
        }

        /// <summary>
        /// Ensures Presentation and SlideMasterIdList exist on the target PresentationPart.
        /// </summary>
        private static void EnsurePresentationAndMasterIdList(PresentationPart targetPresPart)
        {
            if (targetPresPart.Presentation == null)
            {
                targetPresPart.Presentation = new Presentation();
            }

            if (targetPresPart.Presentation.SlideMasterIdList == null)
            {
                targetPresPart.Presentation.SlideMasterIdList = new SlideMasterIdList();
            }
        }

    }
}

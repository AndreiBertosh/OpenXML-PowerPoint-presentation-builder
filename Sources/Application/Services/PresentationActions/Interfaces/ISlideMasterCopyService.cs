using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml.Packaging;

namespace Application.Services.PresentationActions.Interfaces
{
    public interface ISlideMasterCopyService
    {
        Task<string> CopySlideMasterByThemeName(PresentationPart sourcePresPart, PresentationPart targetPresPart, string themeName, bool skipIfExistsInTarget = true);
    }
}

using API.Endpoints.PresentationParser.Models;

using Application.Services.PresentationParser.Models;

using AutoMapper;

namespace API.Endpoints.PresentationParser.Mapping
{
    public class LayoutsInfoMappingProfile : Profile
    {
        public LayoutsInfoMappingProfile()
        {
            CreateMap<RunInfoDomainModel, RunInfoViewModel>();
            CreateMap<ParagraphInfoDomainModel, ParagraphInfoViewModel>();
            CreateMap<ShapeInfoDomainModel, ShapeInfoViewModel>();
            CreateMap<SlideLayoutInfoDomainModel, SlideLayoutInfoViewModel>();
            CreateMap<SlideMasterInfoDomainModel, SlideMasterInfoViewModel>();

            CreateMap<SectionDataDomainModel, SectionDataViewModel>();
            CreateMap<SlideTitleDataDomainModel, SlideTitleDataViewModel>();
            CreateMap<SlideDataDomainModel, SlideDataViewModel>();
        }
    }
}

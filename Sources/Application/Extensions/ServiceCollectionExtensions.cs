using Application.Services.AzureServices;
using Application.Services.AzureServices.Interfaces;
using Application.Services.PresentationActions;
using Application.Services.PresentationActions.Interfaces;
using Application.Services.PresentationData;
using Application.Services.PresentationData.Interfaces;
using Application.Services.PresentationParser;
using Application.Services.PresentationParser.Interfaces;
using Application.Services.SharePointServices;
using Application.Services.SharePointServices.Interfaces;
using Microsoft.Extensions.DependencyInjection;

namespace Application.Extensions
{
    public static class ServiceCollectionExtensions
    {
        public static void RegisterApplicationServices(this IServiceCollection services)
        {
            services.RegisterServices();
        }

        private static void RegisterServices(this IServiceCollection services)
        {
            services.AddScoped<IPresentationParserServices, PresentationParserServices>();
            services.AddScoped<IPresentationDataServices, PresentationDataServices>();
            services.AddScoped<IPresentationActionsServices, PresentationActionsServices>();
            services.AddScoped<IAddNewSlideToPresentationServices, AddNewSlideToPresentationServices>();
            services.AddScoped<ICopySlideServices, CopySlideServices>();
            services.AddScoped<IAzureServices, AzureServices>();
            services.AddScoped<ISharePointServices, SharePointServices>();
            services.AddScoped<ISlideMasterCopyService, SlideMasterCopyService>();
        }
    }
}

using API.Common;
using API.Endpoints.Additional;
using API.Endpoints.Common;
using API.Endpoints.PresentationActions;
using API.Endpoints.PresentationParser;
using API.Endpoints.SharePoint;

using Application.Extensions;

using Infrastructure.Common;
using Infrastructure.Extensions;

using Microsoft.AspNetCore.Http.Features;
using Microsoft.OpenApi.Models;

namespace API
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var builder = WebApplication.CreateBuilder(args);

            // Configure constants layers
            builder.ConfigureConstants();
            builder.ConfigureHealthCheck();
            builder.ConfigureLogging();
            builder.ConfigureGraphAuthProvider();

            builder.Services.AddAutoMapper(
                cfg => { },
                typeof(IAPIMarker).Assembly
            );

            builder.Services.AddControllers();
            // Allow large multipart uploads (up to 2 GB)
            builder.Services.Configure<FormOptions>(options =>
            {
                options.MultipartBodyLengthLimit = 2L * 1024 * 1024 * 1024;
            });

            // Setting up ports for Docker or WSL
            // If the ASPNETCORE_URLS environment variable is not set, use HTTP and HTTPS from appsettings.json manually
            var appSettings = builder.Configuration.GetSection("App").Get<AppSettings>();
            var urls = Environment.GetEnvironmentVariable("ASPNETCORE_URLS") ?? $"http://*:{appSettings.HttpPort};https://*:{appSettings.HttpsPort}";
            builder.WebHost.UseUrls(urls);

            // Reverse proxy configuration
            builder.Services.AddReverseProxy()
                .LoadFromConfig(builder.Configuration.GetSection("ReverseProxy"));

            // Configure infrastructure
            builder.ConfigureInfrastructure();

            // Enable Swagger for API documentation
            builder.Services.AddEndpointsApiExplorer();
            builder.Services.AddSwaggerGen(c =>
            {
                c.SwaggerDoc("v1", new OpenApiInfo { Title = "Presentation builder API", Version = "v1", Description = "API v1 - 001" });
                c.SwaggerDoc("v2", new OpenApiInfo { Title = "Presentation builder API", Version = "v2", Description = "API v2" });
                // Additional configuration if Swagger filters are needed
                // c.OperationFilter<AddApiVersionHeaderFilter>();
            });

            // Register services
            builder.Services.RegisterApplicationServices();

            builder.Services.AddCors(options =>
            {
                options.AddPolicy("AllowAll", policy =>
                {
                    policy.AllowAnyOrigin()
                          .AllowAnyHeader()
                          .AllowAnyMethod();
                });
            });

            // Create the application pipeline
            var app = builder.Build();

            app.MapHealthChecks("/health");

            // Configure Swagger (it will be available at `/`)
            app.UseSwagger();
            app.UseSwaggerUI(c =>
            {
                c.SwaggerEndpoint("/swagger/v1/swagger.json", "v1");
                c.SwaggerEndpoint("/swagger/v2/swagger.json", "v2");
                c.RoutePrefix = string.Empty; // Makes Swagger available at the root level
            });

            // Register versions and routes
            app.RegisterVersionSet();
            app.RegistrationParserEndpoints();
            app.RegistrationActionsEndpoints();
            app.RegistrationAdditionalEndpoints();
            app.RegistrationSharePointEndpoints();

            // HTTPS redirection is enabled only in production
            if (app.Environment.IsProduction())
            {
                app.UseHttpsRedirection();
            }
            else
            {
                Console.WriteLine("HTTPS redirection is disabled for non-production environments.");
            }

            app.UseCors("AllowAll");

            // Enable routing
            app.UseRouting();

            // Connect reverse proxy
            app.MapReverseProxy();
            app.MapControllers();

            // Start the application
            app.Run();
        }
    }
}
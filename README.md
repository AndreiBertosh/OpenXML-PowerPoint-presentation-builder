# Marketing Presentation Builder API

A .NET 9.0 ASP.NET Core API for building and managing marketing presentations. This API provides endpoints for:

- Parsing presentation data from Azure Blob Storage
- Creating and modifying PowerPoint presentations
- Managing slides and templates
- Integrating with SharePoint for file operations
- Downloading and uploading presentation files

The API supports multiple versions (v1 and v2) and includes Swagger documentation for easy exploration. 

# Getting Started

## Prerequisites

- .NET SDK 9.0 or later
- Azure Storage account (for blob storage configuration)
- Azure App credentials (ClientId, TenantId, ClientSecret) for SharePoint integration

## Installation

1. Clone the repository
2. Navigate to the project directory:
   ```bash
   cd marketing_presentation_builder_api/Sources/API
   ```

3. Restore dependencies:
   ```bash
   dotnet restore
   ```

4. Configure `appsettings.json` or `appsettings.Development.json` with your Azure credentials:
   - `AzureApp`: ClientId, TenantId, ClientSecret, SiteHost, SitePath
   - `AzureStorage`: ConnectionString, ContainerName
   - `ApplicationInsights`: ConnectionString (optional)

## How to Start the Project

### Development Mode (with hot reload)

Navigate to the API project directory and run:

```bash
cd Sources/API
dotnet watch run
```

This will start the API server with hot reload enabled, automatically restarting when you make code changes.

### Standard Run Mode

```bash
cd Sources/API
dotnet run
```

### Using a Specific Profile

The project includes launch profiles defined in `Properties/launchSettings.json`:

- **HTTP only**: `dotnet run --launch-profile http`
- **HTTPS**: `dotnet run --launch-profile https`

## Access Points

Once the server is running, you can access:

- **API Base URL**: http://localhost:8080
- **HTTPS URL**: https://localhost:8081
- **Swagger UI**: http://localhost:8080/ (available at root)
- **Health Check**: http://localhost:8080/health

## Troubleshooting

### Port Already in Use

If you encounter "address already in use" error:

```bash
# Find and kill processes using port 8080
lsof -ti:8080 | xargs kill -9

# Or stop all dotnet processes
pkill -f "dotnet.*watch.*run"
```

### Build Errors

If you encounter file lock errors during build:

```bash
# Stop all running dotnet processes
pkill -9 -f "dotnet"

# Clean and rebuild
dotnet clean
dotnet build
```

# Build and Test

## Build

To build the solution:

```bash
dotnet build
```

## Test

TODO: Add test instructions when tests are available 

# Contribute
TODO: Explain how other users and developers can contribute to make your code better. 

If you want to learn more about creating good readme files then refer the following [guidelines](https://docs.microsoft.com/en-us/azure/devops/repos/git/create-a-readme?view=azure-devops). You can also seek inspiration from the below readme files:
- [ASP.NET Core](https://github.com/aspnet/Home)
- [Visual Studio Code](https://github.com/Microsoft/vscode)
- [Chakra Core](https://github.com/Microsoft/ChakraCore)
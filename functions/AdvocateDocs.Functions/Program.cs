using AdvocateDocs.Functions.Auth;
using AdvocateDocs.Functions.Bundles;
using AdvocateDocs.Functions.Functions;
using AdvocateDocs.Functions.Graph;
using Azure.Core;
using Azure.Identity;
using Azure.Storage.Blobs;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

var host = new HostBuilder()
    .ConfigureFunctionsWebApplication(worker => worker.UseMiddleware<EntraAuthMiddleware>())
    .ConfigureServices((context, services) =>
    {
        var config = context.Configuration;
        services.AddApplicationInsightsTelemetryWorkerService();
        services.ConfigureFunctionsApplicationInsights();

        services.AddOptions<EntraOptions>().Bind(config.GetSection(EntraOptions.SectionName)).ValidateOnStart()
            .Validate(o => !string.IsNullOrWhiteSpace(o.TenantId) && !string.IsNullOrWhiteSpace(o.ClientId), "Entra:TenantId and Entra:ClientId are required.");
        services.AddOptions<SharePointOptions>().Bind(config.GetSection(SharePointOptions.SectionName)).ValidateOnStart()
            .Validate(o => !string.IsNullOrWhiteSpace(o.Host), "SharePoint:Host is required.");
        services.AddSingleton<ITokenValidator, TokenValidator>();

        // Managed identity in Azure; developer sign-in (az / Visual Studio) when running locally.
        services.AddSingleton<TokenCredential>(_ => new DefaultAzureCredential(new DefaultAzureCredentialOptions { ExcludeInteractiveBrowserCredential = true }));
        services.AddHttpClient<GraphSiteClient>(c => c.Timeout = TimeSpan.FromMinutes(5));
        services.AddSingleton(_ => new BlobServiceClient(config["AzureWebJobsStorage"]));
        services.AddSingleton<BundleWorkStore>();
    })
    .Build();

await host.RunAsync();

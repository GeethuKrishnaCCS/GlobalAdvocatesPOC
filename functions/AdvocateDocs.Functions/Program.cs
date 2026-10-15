using AdvocateDocs.Functions.Auth;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

var host = new HostBuilder()
    .ConfigureFunctionsWebApplication(worker => worker.UseMiddleware<EntraAuthMiddleware>())
    .ConfigureServices((context, services) =>
    {
        services.AddApplicationInsightsTelemetryWorkerService();
        services.ConfigureFunctionsApplicationInsights();
        services.AddOptions<EntraOptions>().Bind(context.Configuration.GetSection(EntraOptions.SectionName)).ValidateOnStart()
            .Validate(o => !string.IsNullOrWhiteSpace(o.TenantId) && !string.IsNullOrWhiteSpace(o.ClientId), "Entra:TenantId and Entra:ClientId are required.");
        services.AddSingleton<ITokenValidator, TokenValidator>();
    })
    .Build();

await host.RunAsync();

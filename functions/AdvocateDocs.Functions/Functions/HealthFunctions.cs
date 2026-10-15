using AdvocateDocs.Functions.Auth;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;

namespace AdvocateDocs.Functions.Functions;

/// <summary>Liveness and identity probes.</summary>
public sealed class HealthFunctions
{
    private static readonly string Version = typeof(HealthFunctions).Assembly.GetName().Version?.ToString() ?? "unknown";

    /// <summary>Anonymous liveness check.</summary>
    [Function("Health")]
    public IActionResult Health([HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "health")] HttpRequest request)
        => new OkObjectResult(new { status = "ok", version = Version });

    /// <summary>Echoes the authenticated caller — proves the SPFx → Function token path end to end.</summary>
    [Function("Me")]
    public IActionResult Me([HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "me")] HttpRequest request, FunctionContext context)
    {
        var caller = EntraAuthMiddleware.GetCaller(context);
        return new OkObjectResult(new
        {
            name = caller.FindFirst("name")?.Value,
            upn = caller.FindFirst("preferred_username")?.Value ?? caller.FindFirst("upn")?.Value,
            scopes = caller.FindFirst("scp")?.Value
        });
    }
}

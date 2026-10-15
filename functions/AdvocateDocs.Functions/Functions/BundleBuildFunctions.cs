using System.Text.RegularExpressions;
using AdvocateDocs.Functions.Auth;
using AdvocateDocs.Functions.Bundles;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.DurableTask.Client;
using Microsoft.Extensions.Options;

namespace AdvocateDocs.Functions.Functions;

/// <summary>Start a bundle build and poll its progress.</summary>
public sealed partial class BundleBuildFunctions(IOptions<SharePointOptions> sharePoint)
{
    [GeneratedRegex("^[A-Za-z0-9_-]{1,100}$")]
    private static partial Regex ListName();

    /// <summary>POST /api/bundles/build — body: <see cref="BuildRequest"/>. Returns 202 with the instance id.</summary>
    [Function("StartBundleBuild")]
    public async Task<IActionResult> Start(
        [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "bundles/build")] HttpRequest request,
        [DurableClient] DurableTaskClient client,
        FunctionContext context)
    {
        var body = await request.ReadFromJsonAsync<BuildRequest>(context.CancellationToken);
        var problem = Validate(body);
        if (problem is not null)
        {
            return new BadRequestObjectResult(new { error = problem });
        }

        var caller = EntraAuthMiddleware.GetCaller(context);
        var requestedBy = caller.FindFirst("preferred_username")?.Value ?? caller.FindFirst("upn")?.Value ?? "unknown";
        var instanceId = await client.ScheduleNewOrchestrationInstanceAsync(nameof(BuildBundleOrchestrator), new BuildInput(body!, requestedBy));
        return new AcceptedResult($"/api/bundles/build/{instanceId}", new { instanceId });
    }

    /// <summary>GET /api/bundles/build/{instanceId} — status, progress, result or error.</summary>
    [Function("GetBundleBuild")]
    public async Task<IActionResult> Status(
        [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "bundles/build/{instanceId}")] HttpRequest request,
        string instanceId,
        [DurableClient] DurableTaskClient client)
    {
        var instance = await client.GetInstanceAsync(instanceId, getInputsAndOutputs: true);
        if (instance is null)
        {
            return new NotFoundObjectResult(new { error = "Unknown build." });
        }
        return new OkObjectResult(new
        {
            status = instance.RuntimeStatus.ToString(),
            progress = instance.ReadCustomStatusAs<BuildProgress>(),
            result = instance.RuntimeStatus == OrchestrationRuntimeStatus.Completed ? instance.ReadOutputAs<BuildResult>() : null,
            error = instance.FailureDetails?.ErrorMessage
        });
    }

    /// <summary>Only matter sites in this tenant, and list names that are plain URL segments.</summary>
    private string? Validate(BuildRequest? body)
    {
        if (body is null || string.IsNullOrWhiteSpace(body.BundleId) || body.Lists is null)
        {
            return "webUrl, bundleId and lists are required.";
        }
        if (!Uri.TryCreate(body.WebUrl, UriKind.Absolute, out var url) || url.Scheme != Uri.UriSchemeHttps
            || !url.Host.Equals(sharePoint.Value.Host, StringComparison.OrdinalIgnoreCase))
        {
            return "webUrl must be a site in this tenant.";
        }
        var names = new[] { body.Lists.Documents, body.Lists.Bundles, body.Lists.Registry, body.Lists.Items, body.Lists.Manifests };
        return names.All(n => n is not null && ListName().IsMatch(n)) ? null : "List names must be URL names (letters, digits, - and _).";
    }
}

/// <summary>SharePoint tenant the Function serves (app settings section <c>SharePoint</c>).</summary>
public sealed class SharePointOptions
{
    public const string SectionName = "SharePoint";

    /// <summary>e.g. contoso.sharepoint.com — build requests for other hosts are rejected.</summary>
    public string Host { get; init; } = string.Empty;
}

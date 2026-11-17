using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using AdvocateDocs.Functions.Auth;
using AdvocateDocs.Functions.Graph;
using AdvocateDocs.Functions.MatterFolders;
using AdvocateDocs.Functions.SharePoint;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.DurableTask.Client;
using Microsoft.DurableTask.Client.Entities;
using Microsoft.Extensions.Options;

namespace AdvocateDocs.Functions.Functions;

/// <summary>What the add-in puts to pin a matter: its site and the documents library mail is filed to.</summary>
public sealed record MatterFolderRequest(string? WebUrl, string? DocumentsList);

/// <summary>Pinned matters: mail folders in the caller's mailbox, and the timer that files their mail to the matter library.</summary>
public sealed partial class MatterFolderFunctions(MatterFolderService matters, MatterFolderStore store,
    UserGraphTokens tokens, IOptions<SharePointOptions> sharePoint)
{
    /// <summary>Reachable without a user token (see <see cref="EntraAuthMiddleware"/>): Graph is the caller.</summary>
    public const string NotifyFunctionName = "MatterFolderNotify";

    [GeneratedRegex("^[A-Za-z0-9_-]([A-Za-z0-9 _-]{0,98}[A-Za-z0-9_-])?$")]
    private static partial Regex MatterRef();

    private static IActionResult Forbidden(string detail)
        => new ObjectResult(new { error = detail }) { StatusCode = StatusCodes.Status403Forbidden };

    /// <summary>PUT /api/matter-folders/{matterRef} — body: <see cref="MatterFolderRequest"/>. Creates the mail folders; returns their ids.</summary>
    [Function("PutMatterFolders")]
    public async Task<IActionResult> Put(
        [HttpTrigger(AuthorizationLevel.Anonymous, "put", Route = "matter-folders/{matterRef}")] HttpRequest request,
        string matterRef, [DurableClient] DurableTaskClient durable, FunctionContext context)
    {
        var ct = context.CancellationToken;
        var body = await request.ReadFromJsonAsync<MatterFolderRequest>(ct);
        var problem = !MatterRef().IsMatch(matterRef)
            ? "matterRef must be letters, digits, space, - or _ (max 100)."
            : BundleBuildFunctions.ValidateSite(body?.WebUrl, sharePoint.Value.Host) ?? BundleBuildFunctions.ValidateListNames([body?.DocumentsList]);
        if (problem is not null)
        {
            return new BadRequestObjectResult(new { error = problem });
        }

        var caller = EntraAuthMiddleware.GetCaller(context);
        var userId = caller.FindFirst("oid")?.Value;
        var upn = caller.FindFirst("preferred_username")?.Value ?? caller.FindFirst("upn")?.Value;
        var bearer = UserGraphTokens.BearerOf(request.Headers.Authorization);
        if (userId is null || upn is null || bearer is null)
        {
            return Forbidden("The token has no oid or preferred_username/upn claim.");
        }
        try
        {
            // Everything below runs as the caller: no separate access check, SharePoint and Exchange decide.
            await tokens.BeginAsync(userId, bearer, ct);
            var result = await matters.RegisterAsync(userId, upn, matterRef, new Uri(body!.WebUrl!), body.DocumentsList!, ct);
            if (result.IsSuccess)
            {
                // Mail that was already waiting in the folders (re-pin, or moved in while sign-in had expired).
                await durable.Entities.SignalEntityAsync(MatterSweepEntity.IdFor(userId, matterRef), MatterSweepEntity.SweepOperation, cancellation: ct);
            }
            return result.IsSuccess
                ? new OkObjectResult(new { matterRef, folders = result.Value!.Folders.Select(f => new { name = f.Name, mailFolderId = f.MailFolderId }) })
                : new NotFoundObjectResult(new { error = result.Error });
        }
        catch (SignInRequiredException ex)
        {
            return Forbidden($"{ex.Message} ({ex.InnerException?.Message})");
        }
        catch (GraphRequestException ex) when (ex.Status is HttpStatusCode.Forbidden or HttpStatusCode.NotFound)
        {
            return Forbidden($"You don't have access to this matter or mailbox. {ex.Message}");
        }
    }

    /// <summary>DELETE /api/matter-folders/{matterRef} — unpins; deletes the matter's empty mail folders. 204 also when nothing was pinned.</summary>
    [Function("DeleteMatterFolders")]
    public async Task<IActionResult> Delete(
        [HttpTrigger(AuthorizationLevel.Anonymous, "delete", Route = "matter-folders/{matterRef}")] HttpRequest request,
        string matterRef, FunctionContext context)
    {
        var ct = context.CancellationToken;
        var userId = EntraAuthMiddleware.GetCaller(context).FindFirst("oid")?.Value;
        var bearer = UserGraphTokens.BearerOf(request.Headers.Authorization);
        if (userId is null || bearer is null || !MatterRef().IsMatch(matterRef))
        {
            return new BadRequestObjectResult(new { error = "A valid matterRef and an oid claim are required." });
        }
        try
        {
            await tokens.BeginAsync(userId, bearer, ct);
            await matters.UnregisterAsync(userId, matterRef, ct);
        }
        catch (SignInRequiredException ex)
        {
            return Forbidden($"{ex.Message} ({ex.InnerException?.Message})");
        }
        // Nothing left to file for this user: drop the stored sign-in.
        await using var remaining = store.ListAsync(userId, ct).GetAsyncEnumerator(ct);
        if (!await remaining.MoveNextAsync())
        {
            await tokens.EndAsync(userId, ct);
        }
        return new NoContentResult();
    }

    /// <summary>GET /api/matter-folders — the caller's pinned matters with filing status.</summary>
    [Function("ListMatterFolders")]
    public async Task<IActionResult> List(
        [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "matter-folders")] HttpRequest request, FunctionContext context)
    {
        var userId = EntraAuthMiddleware.GetCaller(context).FindFirst("oid")?.Value;
        if (userId is null)
        {
            return new BadRequestObjectResult(new { error = "The token has no oid claim." });
        }
        var pinned = new List<object>();
        await foreach (var stored in store.ListAsync(userId, context.CancellationToken))
        {
            var s = stored.State;
            pinned.Add(new { matterRef = s.MatterRef, filedCount = s.FiledCount, lastSweepUtc = s.LastSweepUtc, lastError = s.LastError });
        }
        return new OkObjectResult(new { matters = pinned });
    }

    /// <summary>
    /// POST /api/matter-folders/notify?u={oid}&amp;m={matterRef} — called by Microsoft Graph, without a user token, when mail
    /// arrives in a pinned matter folder. Answers the subscription handshake; otherwise queues a sweep of that matter when
    /// the notification carries the matter's secret. Always 202 quickly: Graph drops slow endpoints.
    /// </summary>
    [Function(NotifyFunctionName)]
    public async Task<IActionResult> Notify(
        [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = MatterFolderService.NotifyRoute)] HttpRequest request,
        [DurableClient] DurableTaskClient durable, FunctionContext context)
    {
        var ct = context.CancellationToken;
        if (request.Query.TryGetValue("validationToken", out var validationToken))
        {
            return new ContentResult { Content = validationToken.ToString(), ContentType = "text/plain", StatusCode = StatusCodes.Status200OK };
        }
        string? userId = request.Query["u"], matterRef = request.Query["m"];
        if (!Guid.TryParse(userId, out _) || matterRef is null || !MatterRef().IsMatch(matterRef))
        {
            return new AcceptedResult();
        }
        var notifications = (await request.ReadFromJsonAsync<JsonObject>(ct))?["value"] as JsonArray;
        var secret = (await store.GetAsync(userId, matterRef, ct))?.State.NotificationSecret;
        if (secret is not null && notifications is not null && notifications.Any(n => IsSecret(n?["clientState"]?.GetValue<string>(), secret)))
        {
            await durable.Entities.SignalEntityAsync(MatterSweepEntity.IdFor(userId, matterRef), MatterSweepEntity.SweepOperation, cancellation: ct);
        }
        return new AcceptedResult();
    }

    /// <summary>Once a day: renews every pinned matter's notifications and sweeps it, in case a notification was missed.</summary>
    [Function("MatterFolderMaintenance")]
    public async Task Maintain([TimerTrigger("0 30 1 * * *")] TimerInfo timer, [DurableClient] DurableTaskClient durable, FunctionContext context)
    {
        await foreach (var stored in store.ListAsync(null, context.CancellationToken))
        {
            await durable.Entities.SignalEntityAsync(MatterSweepEntity.IdFor(stored.State.UserId, stored.State.MatterRef),
                MatterSweepEntity.MaintainOperation, cancellation: context.CancellationToken);
        }
    }

    private static bool IsSecret(string? candidate, string secret)
        => candidate is not null && CryptographicOperations.FixedTimeEquals(Encoding.UTF8.GetBytes(candidate), Encoding.UTF8.GetBytes(secret));
}

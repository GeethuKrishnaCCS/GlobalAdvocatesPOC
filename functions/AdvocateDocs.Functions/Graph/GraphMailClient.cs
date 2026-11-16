using System.Globalization;
using System.Net;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json.Nodes;
using Azure.Core;

namespace AdvocateDocs.Functions.Graph;

/// <summary>A mail folder's message counts.</summary>
public sealed record MailFolderInfo(string Id, int TotalItemCount, int ChildFolderCount);

/// <summary>The message fields the sweep needs.</summary>
public sealed record MailMessageRef(string Id, string? Subject, DateTimeOffset? ReceivedUtc, IReadOnlyList<string> Categories);

/// <summary>Messages changed since the last delta link, and the link for the next round (null when Graph returned none).</summary>
public sealed record MailDelta(IReadOnlyList<MailMessageRef> Messages, string? DeltaLink);

/// <summary>
/// Minimal Microsoft Graph client for a user's mail folders and messages, authenticated as that user
/// (delegated Mail.ReadWrite through <see cref="UserGraph"/>). Respects Retry-After once on 429/503, then gives up.
/// </summary>
public sealed class GraphMailClient(HttpClient http, TokenCredential credential)
{
    private const string GraphRoot = "https://graph.microsoft.com/v1.0";
    private static readonly string[] Scopes = ["https://graph.microsoft.com/.default"];
    private const string DeltaSelect = "id,subject,receivedDateTime,categories";
    private static readonly TimeSpan DefaultRetryDelay = TimeSpan.FromSeconds(5);
    private static readonly TimeSpan MaxRetryDelay = TimeSpan.FromSeconds(30);

    /// <summary>Id of the folder named <paramref name="name"/> under <paramref name="parentId"/> (null = top level), or null.</summary>
    public async Task<string?> FindFolderAsync(string userId, string? parentId, string name, CancellationToken ct)
    {
        var filter = Uri.EscapeDataString($"displayName eq '{name.Replace("'", "''")}'");
        var page = await GetJsonAsync($"{FoldersUrl(userId, parentId)}?$filter={filter}&$select=id", ct);
        return page["value"]!.AsArray().FirstOrDefault()?["id"]?.GetValue<string>();
    }

    /// <summary>Creates a folder under <paramref name="parentId"/> (null = top level, beside Inbox) and returns its id.</summary>
    public async Task<string> CreateFolderAsync(string userId, string? parentId, string name, CancellationToken ct)
    {
        using var response = await SendAsync(HttpMethod.Post, FoldersUrl(userId, parentId), new { displayName = name }, ct);
        return (await ReadJsonAsync(response, ct))["id"]!.GetValue<string>();
    }

    /// <summary>The folder's message counts, or null when it no longer exists.</summary>
    public async Task<MailFolderInfo?> GetFolderAsync(string userId, string folderId, CancellationToken ct)
    {
        using var response = await SendAsync(HttpMethod.Get, $"{GraphRoot}/users/{userId}/mailFolders/{folderId}?$select=id,totalItemCount,childFolderCount", null, ct);
        if (response.StatusCode == HttpStatusCode.NotFound)
        {
            return null;
        }
        var folder = await ReadJsonAsync(response, ct);
        return new MailFolderInfo(folderId, folder["totalItemCount"]!.GetValue<int>(), folder["childFolderCount"]!.GetValue<int>());
    }

    /// <summary>Deletes a mail folder; a missing folder is not an error.</summary>
    public async Task DeleteFolderAsync(string userId, string folderId, CancellationToken ct)
    {
        using var response = await SendAsync(HttpMethod.Delete, $"{GraphRoot}/users/{userId}/mailFolders/{folderId}", null, ct);
        if (response.StatusCode != HttpStatusCode.NotFound)
        {
            await GraphSiteClient.EnsureSuccessAsync(response, "delete mail folder");
        }
    }

    /// <summary>
    /// All message changes since <paramref name="deltaLink"/> (null = from the start); removed entries are ignored.
    /// Throws <see cref="GraphRequestException"/> (404) when the folder no longer exists.
    /// </summary>
    public async Task<MailDelta> GetDeltaAsync(string userId, string folderId, string? deltaLink, CancellationToken ct)
    {
        // Stored links carry our bearer token to wherever they point: only follow Graph's own.
        var next = deltaLink is not null && deltaLink.StartsWith(GraphRoot, StringComparison.Ordinal)
            ? deltaLink
            : $"{GraphRoot}/users/{userId}/mailFolders/{folderId}/messages/delta?$select={DeltaSelect}";
        var messages = new List<MailMessageRef>();
        string? finalLink = null;
        while (next is not null)
        {
            var page = await GetJsonAsync(next, ct);
            messages.AddRange(page["value"]!.AsArray().Where(n => n!["@removed"] is null).Select(n => ToMessage(n!.AsObject())));
            next = page["@odata.nextLink"]?.GetValue<string>();
            finalLink = page["@odata.deltaLink"]?.GetValue<string>() ?? finalLink;
        }
        return new MailDelta(messages, finalLink);
    }

    /// <summary>The message as MIME (.eml) bytes.</summary>
    public async Task<byte[]> DownloadMimeAsync(string userId, string messageId, CancellationToken ct)
    {
        using var response = await SendAsync(HttpMethod.Get, $"{GraphRoot}/users/{userId}/messages/{messageId}/$value", null, ct);
        await GraphSiteClient.EnsureSuccessAsync(response, "download message");
        return await response.Content.ReadAsByteArrayAsync(ct);
    }

    /// <summary>Sets the message's categories (send the full list: Graph replaces it).</summary>
    public async Task SetCategoriesAsync(string userId, string messageId, IEnumerable<string> categories, CancellationToken ct)
    {
        using var response = await SendAsync(HttpMethod.Patch, $"{GraphRoot}/users/{userId}/messages/{messageId}", new { categories }, ct);
        await GraphSiteClient.EnsureSuccessAsync(response, "tag message");
    }

    /// <summary>Subscribes to messages arriving in a mail folder (created; a move into the folder counts) and returns the subscription id.</summary>
    public async Task<string> CreateSubscriptionAsync(string userId, string folderId, Uri notificationUrl, string clientState,
        DateTimeOffset expires, CancellationToken ct)
    {
        using var response = await SendAsync(HttpMethod.Post, $"{GraphRoot}/subscriptions", new
        {
            changeType = "created", notificationUrl = notificationUrl.AbsoluteUri, clientState, // AbsoluteUri keeps %20; ToString() would unescape it
            resource = $"users/{userId}/mailFolders('{folderId}')/messages", expirationDateTime = expires.ToUniversalTime()
        }, ct);
        return (await ReadJsonAsync(response, ct))["id"]!.GetValue<string>();
    }

    /// <summary>Extends a subscription; false when it no longer exists.</summary>
    public async Task<bool> RenewSubscriptionAsync(string subscriptionId, DateTimeOffset expires, CancellationToken ct)
    {
        using var response = await SendAsync(HttpMethod.Patch, $"{GraphRoot}/subscriptions/{subscriptionId}", new { expirationDateTime = expires.ToUniversalTime() }, ct);
        if (response.StatusCode == HttpStatusCode.NotFound)
        {
            return false;
        }
        await GraphSiteClient.EnsureSuccessAsync(response, "renew mail notification");
        return true;
    }

    /// <summary>Deletes a subscription; a missing one is not an error.</summary>
    public async Task DeleteSubscriptionAsync(string subscriptionId, CancellationToken ct)
    {
        using var response = await SendAsync(HttpMethod.Delete, $"{GraphRoot}/subscriptions/{subscriptionId}", null, ct);
        if (response.StatusCode != HttpStatusCode.NotFound)
        {
            await GraphSiteClient.EnsureSuccessAsync(response, "delete mail notification");
        }
    }

    private static string FoldersUrl(string userId, string? parentId)
        => parentId is null ? $"{GraphRoot}/users/{userId}/mailFolders" : $"{GraphRoot}/users/{userId}/mailFolders/{parentId}/childFolders";

    private static MailMessageRef ToMessage(JsonObject node)
        => new(node["id"]!.GetValue<string>(), node["subject"]?.GetValue<string>(),
            node["receivedDateTime"] is { } received ? DateTimeOffset.Parse(received.GetValue<string>(), CultureInfo.InvariantCulture) : null,
            node["categories"]?.AsArray().Select(c => c!.GetValue<string>()).ToList() ?? []);

    private async Task<JsonObject> GetJsonAsync(string url, CancellationToken ct)
    {
        using var response = await SendAsync(HttpMethod.Get, url, null, ct);
        return await ReadJsonAsync(response, ct);
    }

    private static async Task<JsonObject> ReadJsonAsync(HttpResponseMessage response, CancellationToken ct)
    {
        await GraphSiteClient.EnsureSuccessAsync(response, "call Graph mail");
        return (await response.Content.ReadFromJsonAsync<JsonObject>(ct))!;
    }

    /// <summary>Sends with a fresh token; on 429/503 waits Retry-After (capped) and retries once.</summary>
    private async Task<HttpResponseMessage> SendAsync(HttpMethod method, string url, object? body, CancellationToken ct)
    {
        for (var attempt = 1; ; attempt++)
        {
            var token = await credential.GetTokenAsync(new TokenRequestContext(Scopes), ct);
            using var request = new HttpRequestMessage(method, url) { Content = body is null ? null : JsonContent.Create(body) };
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
            var response = await http.SendAsync(request, ct);
            if (attempt > 1 || response.StatusCode is not (HttpStatusCode.TooManyRequests or HttpStatusCode.ServiceUnavailable))
            {
                return response;
            }
            var delay = response.Headers.RetryAfter?.Delta ?? DefaultRetryDelay;
            response.Dispose();
            await Task.Delay(delay < MaxRetryDelay ? delay : MaxRetryDelay, ct);
        }
    }
}

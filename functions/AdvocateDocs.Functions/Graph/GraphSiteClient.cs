using System.Net;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;
using System.Text.Json.Nodes;
using Azure.Core;

namespace AdvocateDocs.Functions.Graph;

/// <summary>Drive item reference returned by Graph.</summary>
public sealed record DriveItemRef(string DriveId, string ItemId, string Name, string WebUrl);

/// <summary>
/// Minimal Microsoft Graph client for SharePoint sites, lists and files, authenticated as the Function's own
/// identity (managed identity in Azure; developer sign-in locally). Retries throttled calls (429/503).
/// </summary>
public sealed class GraphSiteClient(HttpClient http, TokenCredential credential)
{
    private const string GraphRoot = "https://graph.microsoft.com/v1.0";
    private static readonly string[] Scopes = ["https://graph.microsoft.com/.default"];
    private const int MaxAttempts = 5;
    /// <summary>SharePoint field names are case-sensitive: send property names as written (JsonContent defaults to camelCase).</summary>
    private static readonly JsonSerializerOptions AsWritten = new();
    /// <summary>Upload session chunk: must be a multiple of 320 KiB.</summary>
    private const int UploadChunkSize = 320 * 1024 * 16;

    public async Task<string> GetSiteIdAsync(Uri webUrl, CancellationToken ct)
    {
        var site = await GetJsonAsync($"/sites/{webUrl.Host}:{webUrl.AbsolutePath.TrimEnd('/')}?$select=id", ct);
        return site["id"]!.GetValue<string>();
    }

    /// <summary>List ids keyed by list URL name (Graph "name"), never by display title.</summary>
    public async Task<IReadOnlyDictionary<string, string>> GetListIdsByNameAsync(string siteId, CancellationToken ct)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        await foreach (var list in GetPagedAsync($"/sites/{siteId}/lists?$select=id,name", ct))
        {
            result[list["name"]!.GetValue<string>()] = list["id"]!.GetValue<string>();
        }
        return result;
    }

    public async Task<string> GetListDriveIdAsync(string siteId, string listId, CancellationToken ct)
        => (await GetJsonAsync($"/sites/{siteId}/lists/{listId}/drive?$select=id", ct))["id"]!.GetValue<string>();

    /// <summary>List items whose (indexed) text field equals a value, with the requested fields expanded.</summary>
    public async Task<IReadOnlyList<JsonObject>> GetItemsWhereAsync(string siteId, string listId, string field, string value,
        IEnumerable<string> select, CancellationToken ct)
    {
        var url = $"/sites/{siteId}/lists/{listId}/items?$expand=fields($select={string.Join(',', select)})"
            + $"&$filter=fields/{field} eq '{value.Replace("'", "''")}'&$top=999";
        var items = new List<JsonObject>();
        await foreach (var item in GetPagedAsync(url, ct))
        {
            items.Add(item);
        }
        return items;
    }

    public async Task<DriveItemRef> GetListItemDriveItemAsync(string siteId, string listId, int itemId, CancellationToken ct)
    {
        var item = await GetJsonAsync($"/sites/{siteId}/lists/{listId}/items/{itemId}/driveItem?$select=id,name,webUrl,parentReference", ct);
        return new DriveItemRef(item["parentReference"]!["driveId"]!.GetValue<string>(), item["id"]!.GetValue<string>(),
            item["name"]!.GetValue<string>(), item["webUrl"]!.GetValue<string>());
    }

    /// <summary>File content; with <paramref name="asPdf"/>, Graph's conversion. Returns null when Graph can't convert the type.</summary>
    public async Task<byte[]?> DownloadAsync(string driveId, string itemId, bool asPdf, CancellationToken ct)
    {
        using var response = await SendAsync(() => new HttpRequestMessage(HttpMethod.Get,
            $"{GraphRoot}/drives/{driveId}/items/{itemId}/content{(asPdf ? "?format=pdf" : string.Empty)}"), ct);
        if (asPdf && response.StatusCode is HttpStatusCode.NotAcceptable or HttpStatusCode.BadRequest or HttpStatusCode.UnsupportedMediaType)
        {
            return null;
        }
        await EnsureSuccessAsync(response, "download file");
        return await response.Content.ReadAsByteArrayAsync(ct);
    }

    /// <summary>SharePoint's rendered preview image of a file (first page only); null when it has none.</summary>
    public async Task<byte[]?> DownloadPreviewImageAsync(string driveId, string itemId, CancellationToken ct)
    {
        using var response = await SendAsync(() => new HttpRequestMessage(HttpMethod.Get,
            $"{GraphRoot}/drives/{driveId}/items/{itemId}/thumbnails/0/c1600x1600/content"), ct);
        if (response.StatusCode is HttpStatusCode.NotFound or HttpStatusCode.BadRequest)
        {
            return null;
        }
        await EnsureSuccessAsync(response, "download preview image");
        return await response.Content.ReadAsByteArrayAsync(ct);
    }

    /// <summary>Uploads (replacing, which adds a version) a file of any size to a drive path; returns the drive item and its list item id.</summary>
    public async Task<(DriveItemRef Item, int ListItemId)> UploadAsync(string driveId, string path, byte[] content, CancellationToken ct)
    {
        var session = await PostJsonAsync($"/drives/{driveId}/root:/{EscapePath(path)}:/createUploadSession",
            new { item = new Dictionary<string, string> { ["@microsoft.graph.conflictBehavior"] = "replace" } }, ct);
        var uploadUrl = session["uploadUrl"]!.GetValue<string>();
        JsonObject? item = null;
        for (var start = 0; start < content.Length; start += UploadChunkSize)
        {
            var length = Math.Min(UploadChunkSize, content.Length - start);
            // Pre-authorised URL: no bearer token on chunk uploads.
            using var request = new HttpRequestMessage(HttpMethod.Put, uploadUrl) { Content = new ByteArrayContent(content, start, length) };
            request.Content.Headers.ContentRange = new ContentRangeHeaderValue(start, start + length - 1, content.Length);
            using var response = await http.SendAsync(request, ct);
            await EnsureSuccessAsync(response, "upload bundle chunk");
            if (response.StatusCode is HttpStatusCode.OK or HttpStatusCode.Created)
            {
                item = await response.Content.ReadFromJsonAsync<JsonObject>(ct);
            }
        }
        var itemId = item!["id"]!.GetValue<string>();
        var withListItem = await GetJsonAsync($"/drives/{driveId}/items/{itemId}?$select=id,name,webUrl&$expand=listItem($select=id)", ct);
        return (new DriveItemRef(driveId, itemId, withListItem["name"]!.GetValue<string>(), withListItem["webUrl"]!.GetValue<string>()),
            int.Parse(withListItem["listItem"]!["id"]!.GetValue<string>()));
    }

    /// <summary>Deletes a drive item (to the recycle bin); a missing item is not an error.</summary>
    public async Task DeleteAsync(string driveId, string itemId, CancellationToken ct)
    {
        using var response = await SendAsync(() => new HttpRequestMessage(HttpMethod.Delete, $"{GraphRoot}/drives/{driveId}/items/{itemId}"), ct);
        if (response.StatusCode != HttpStatusCode.NotFound)
        {
            await EnsureSuccessAsync(response, "delete temporary copy");
        }
    }

    public async Task<JsonObject> GetListItemFieldsAsync(string siteId, string listId, int itemId, CancellationToken ct)
        => (await GetJsonAsync($"/sites/{siteId}/lists/{listId}/items/{itemId}/fields", ct));

    public async Task UpdateListItemFieldsAsync(string siteId, string listId, int itemId, object fields, CancellationToken ct)
    {
        using var response = await SendAsync(() => new HttpRequestMessage(HttpMethod.Patch, $"{GraphRoot}/sites/{siteId}/lists/{listId}/items/{itemId}/fields")
        {
            Content = JsonContent.Create(fields, options: AsWritten)
        }, ct);
        await EnsureSuccessAsync(response, "update list item");
    }

    private async IAsyncEnumerable<JsonObject> GetPagedAsync(string relativeUrl, [System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken ct)
    {
        string? next = GraphRoot + relativeUrl;
        while (next is not null)
        {
            var page = await GetJsonAbsoluteAsync(next, ct);
            foreach (var node in page["value"]!.AsArray())
            {
                yield return node!.AsObject();
            }
            next = page["@odata.nextLink"]?.GetValue<string>();
        }
    }

    private Task<JsonObject> GetJsonAsync(string relativeUrl, CancellationToken ct) => GetJsonAbsoluteAsync(GraphRoot + relativeUrl, ct);

    private async Task<JsonObject> GetJsonAbsoluteAsync(string url, CancellationToken ct)
    {
        using var response = await SendAsync(() =>
        {
            var request = new HttpRequestMessage(HttpMethod.Get, url);
            // Filtering list items on an indexed column is fine even past the list view threshold.
            request.Headers.Add("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly");
            return request;
        }, ct);
        await EnsureSuccessAsync(response, "read from Graph");
        return (await response.Content.ReadFromJsonAsync<JsonObject>(ct))!;
    }

    private async Task<JsonObject> PostJsonAsync(string relativeUrl, object body, CancellationToken ct)
    {
        using var response = await SendAsync(() => new HttpRequestMessage(HttpMethod.Post, GraphRoot + relativeUrl) { Content = JsonContent.Create(body, options: AsWritten) }, ct);
        await EnsureSuccessAsync(response, "write to Graph");
        return (await response.Content.ReadFromJsonAsync<JsonObject>(ct))!;
    }

    /// <summary>Sends with a fresh token; retries 429/503 honouring Retry-After.</summary>
    private async Task<HttpResponseMessage> SendAsync(Func<HttpRequestMessage> create, CancellationToken ct)
    {
        for (var attempt = 1; ; attempt++)
        {
            var token = await credential.GetTokenAsync(new TokenRequestContext(Scopes), ct);
            var request = create();
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
            var response = await http.SendAsync(request, ct);
            if (attempt >= MaxAttempts || response.StatusCode is not (HttpStatusCode.TooManyRequests or HttpStatusCode.ServiceUnavailable))
            {
                return response;
            }
            var delay = response.Headers.RetryAfter?.Delta ?? TimeSpan.FromSeconds(Math.Pow(2, attempt));
            response.Dispose();
            await Task.Delay(delay, ct);
        }
    }

    private static async Task EnsureSuccessAsync(HttpResponseMessage response, string action)
    {
        if (response.IsSuccessStatusCode)
        {
            return;
        }
        var body = await response.Content.ReadAsStringAsync();
        string? message = null;
        try { message = JsonNode.Parse(body)?["error"]?["message"]?.GetValue<string>(); } catch (JsonException) { /* non-JSON error body */ }
        throw new GraphRequestException($"Couldn't {action} (HTTP {(int)response.StatusCode}): {message ?? response.ReasonPhrase}", response.StatusCode);
    }

    private static string EscapePath(string path) => string.Join('/', path.Split('/').Select(Uri.EscapeDataString));
}

/// <summary>A Graph call failed with a non-success status.</summary>
public sealed class GraphRequestException(string message, HttpStatusCode status) : Exception(message)
{
    public HttpStatusCode Status { get; } = status;
}

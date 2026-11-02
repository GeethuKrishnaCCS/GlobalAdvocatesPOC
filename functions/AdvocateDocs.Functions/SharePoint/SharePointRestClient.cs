using System.Net;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json.Nodes;
using AdvocateDocs.Functions.Functions;
using Azure.Core;
using Microsoft.Extensions.Options;

namespace AdvocateDocs.Functions.SharePoint;

/// <summary>
/// SharePoint REST calls Graph can't make, authenticated as the Function's own identity (needs the
/// Office 365 SharePoint Online Sites.Selected app role plus the site grant). Retries throttled calls (429/503).
/// </summary>
public sealed class SharePointRestClient(HttpClient http, TokenCredential credential, IOptions<SharePointOptions> options)
{
    private const int MaxAttempts = 5;
    /// <summary>SharePoint REST property names are case-sensitive (FieldName, not fieldName); JsonContent defaults to camelCase.</summary>
    private static readonly System.Text.Json.JsonSerializerOptions AsWritten = new();

    /// <summary>
    /// Sets list item columns without adding a version (ValidateUpdateListItem, bNewDocumentUpdate=true).
    /// Values use SharePoint's form formats (numbers and choices as plain strings).
    /// </summary>
    public async Task UpdateFieldsWithoutVersionAsync(Uri webUrl, string listId, int itemId,
        IReadOnlyDictionary<string, string> values, CancellationToken ct)
    {
        var url = $"{webUrl.GetLeftPart(UriPartial.Path).TrimEnd('/')}/_api/web/lists(guid'{listId}')/items({itemId})/ValidateUpdateListItem";
        var body = new
        {
            formValues = values.Select(v => new { FieldName = v.Key, FieldValue = v.Value }),
            bNewDocumentUpdate = true
        };
        using var response = await SendAsync(() => new HttpRequestMessage(HttpMethod.Post, url) { Content = JsonContent.Create(body, options: AsWritten) }, ct);
        var json = await response.Content.ReadAsStringAsync(ct);
        if (!response.IsSuccessStatusCode)
        {
            throw new HttpRequestException($"Couldn't update list item (HTTP {(int)response.StatusCode}): {ErrorMessage(json) ?? response.ReasonPhrase}",
                null, response.StatusCode);
        }
        // A 200 can still carry per-field validation failures.
        var failed = JsonNode.Parse(json)?["value"]?.AsArray()
            .Where(f => f?["HasException"]?.GetValue<bool>() == true)
            .Select(f => $"{f!["FieldName"]}: {f["ErrorMessage"]}")
            .ToList();
        if (failed is { Count: > 0 })
        {
            throw new InvalidOperationException($"Couldn't set columns: {string.Join("; ", failed)}");
        }
    }

    private async Task<HttpResponseMessage> SendAsync(Func<HttpRequestMessage> create, CancellationToken ct)
    {
        var scopes = new[] { $"https://{options.Value.Host}/.default" };
        for (var attempt = 1; ; attempt++)
        {
            var token = await credential.GetTokenAsync(new TokenRequestContext(scopes), ct);
            var request = create();
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
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

    private static string? ErrorMessage(string body)
    {
        try
        {
            var root = JsonNode.Parse(body);
            var message = (root?["odata.error"] ?? root?["error"])?["message"];
            // verbose/minimal metadata nest the text under "value"; nometadata returns it directly.
            return message is JsonObject nested ? nested["value"]?.GetValue<string>() : message?.GetValue<string>();
        }
        catch (Exception)
        {
            return null;
        }
    }
}

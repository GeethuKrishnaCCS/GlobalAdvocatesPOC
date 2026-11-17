using System.Text.Json;
using AdvocateDocs.Functions.Graph;
using AdvocateDocs.Functions.SharePoint;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;

namespace AdvocateDocs.Functions.Bundles;

/// <summary>The bundle build's steps. Each is idempotent so Durable Functions can retry it.</summary>
public sealed class BundleActivities(GraphSiteClient graph, SharePointRestClient sharePoint, MemberConverter converter, BundleWorkStore work, ILogger<BundleActivities> logger)
{
    private static readonly string[] MemberFields = ["SortOrder", "SourceUniqueId", "SourceItemId", "IndexTitle"];
    private static readonly JsonSerializerOptions ManifestJson = new() { WriteIndented = true, PropertyNamingPolicy = JsonNamingPolicy.CamelCase };

    /// <summary>Resolves site and list ids and loads the bundle's members in order.</summary>
    [Function(nameof(LoadPlan))]
    public async Task<BundlePlan> LoadPlan([ActivityTrigger] BuildInput input, CancellationToken ct)
    {
        var request = input.Request;
        var siteId = await graph.GetSiteIdAsync(new Uri(request.WebUrl), ct);
        var lists = await graph.GetListIdsByNameAsync(siteId, ct);
        string ListId(string name) => lists.TryGetValue(name, out var id) ? id : throw new InvalidOperationException($"List '{name}' not found on the matter site.");

        var registryId = ListId(request.Lists.Registry);
        var header = (await graph.GetItemsWhereAsync(siteId, registryId, "BundleId", request.BundleId, ["Title", "BundleId"], ct)).SingleOrDefault()
            ?? throw new InvalidOperationException($"Bundle {request.BundleId} not found.");
        var rows = await graph.GetItemsWhereAsync(siteId, ListId(request.Lists.Items), "BundleId", request.BundleId, MemberFields, ct);

        var members = rows
            .Select(r => r["fields"]!)
            .OrderBy(f => f["SortOrder"]?.GetValue<double>() ?? double.MaxValue)
            .Select((f, i) => new PlannedMember(
                i + 1,
                (int)(f["SourceItemId"]?.GetValue<double>() ?? 0),
                f["SourceUniqueId"]?.GetValue<string>() ?? string.Empty,
                f["IndexTitle"]?.GetValue<string>() ?? $"Document {i + 1}"))
            .ToList();

        // Optional: without it, email attachments Graph would have to convert become placeholders.
        string? conversionId = null;
        if (request.Lists.Conversion is { } conversion && !lists.TryGetValue(conversion, out conversionId))
        {
            logger.LogWarning("Conversion library {Name} not found; Office-type email attachments will be placeholders", conversion);
        }

        return new BundlePlan(siteId, ListId(request.Lists.Documents), ListId(request.Lists.Bundles), registryId, ListId(request.Lists.Manifests),
            int.Parse(header["id"]!.GetValue<string>()), request.BundleId, header["fields"]!["Title"]!.GetValue<string>(), members, conversionId);
    }

    /// <summary>Turns one member into PDF (as-is, Graph conversion, text rendering or a placeholder) and parks it in blob storage.</summary>
    [Function(nameof(ConvertMember))]
    public async Task<ConvertedMember> ConvertMember([ActivityTrigger] ConvertInput input, CancellationToken ct)
    {
        var m = input.Member;
        var blobName = $"{input.WorkPrefix}/members/{m.Order:D5}.pdf";
        var (pdf, kind, note, attachments) = await converter.ConvertAsync(input, ct);
        await work.PutAsync(blobName, pdf, ct);

        var parts = new List<ConvertedPart>(attachments.Count);
        for (var i = 0; i < attachments.Count; i++)
        {
            var a = attachments[i];
            var partBlob = $"{input.WorkPrefix}/members/{m.Order:D5}-a{i + 1:D3}.pdf";
            await work.PutAsync(partBlob, a.Pdf, ct);
            parts.Add(new ConvertedPart(a.Label, a.Level, a.Title, a.Kind, partBlob, a.Note));
        }
        return new ConvertedMember(m.Order, m.Title, kind, blobName, m.SourceUniqueId, m.SourceItemId, note, parts);
    }

    /// <summary>Builds the bundle PDF (index, stamps, bookmarks) from the parked members.</summary>
    [Function(nameof(AssembleBundle))]
    public async Task<AssembleOutput> AssembleBundle([ActivityTrigger] AssembleInput input, CancellationToken ct)
    {
        var members = new List<MemberPdf>(input.Members.Count);
        foreach (var m in input.Members)
        {
            members.Add(new MemberPdf(m.Order, m.Title, m.Kind, await work.GetAsync(m.BlobName, ct), m.Note));
            foreach (var a in m.Attachments ?? [])
            {
                members.Add(new MemberPdf(m.Order, a.Title, a.Kind, await work.GetAsync(a.BlobName, ct), a.Note, a.Label, a.Level));
            }
        }
        var bundle = PdfBundleAssembler.Assemble(input.Title, members);
        var blobName = $"{input.WorkPrefix}/bundle.pdf";
        await work.PutAsync(blobName, bundle.Pdf, ct);
        return new AssembleOutput(blobName, bundle.PageCount, bundle.IndexPageCount, bundle.Entries);
    }

    /// <summary>Uploads the PDF as a new version in the Bundles library, writes the manifest, marks the bundle Built, cleans up.</summary>
    [Function(nameof(PublishBundle))]
    public async Task<BuildResult> PublishBundle([ActivityTrigger] PublishInput input, CancellationToken ct)
    {
        var plan = input.Plan;
        var pdf = await work.GetAsync(input.Assembled.BlobName, ct);

        var bundlesDrive = await graph.GetListDriveIdAsync(plan.SiteId, plan.BundlesListId, ct);
        var (file, fileItemId) = await graph.UploadAsync(bundlesDrive, $"{SafeFileName(plan.Title)}.pdf", pdf, ct);
        var fields = await graph.GetListItemFieldsAsync(plan.SiteId, plan.BundlesListId, fileItemId, ct);
        var version = (int)(fields["ManifestVersion"]?.GetValue<double>() ?? 0) + 1;
        var webUrl = new Uri(input.Build.Request.WebUrl);
        // Columns go on the version the upload just created, so each build adds exactly one version.
        await SetColumnsAsync(plan.SiteId, webUrl, plan.BundlesListId, fileItemId, new Dictionary<string, object>
        {
            ["BundleId"] = plan.BundleId, ["ManifestVersion"] = version, ["BuildStatus"] = "Ready",
            ["PageCount"] = input.Assembled.PageCount, ["VolumeNo"] = 1
        }, ct);

        var bySource = input.Members.ToDictionary(m => m.Order);
        var manifest = new
        {
            plan.BundleId, plan.Title, Version = version, BuiltAtUtc = DateTime.UtcNow, BuiltBy = input.Build.RequestedBy,
            input.Assembled.PageCount, input.Assembled.IndexPageCount, PdfUrl = file.WebUrl,
            Entries = input.Assembled.Entries.Select(e => new
            {
                e.Order, e.Label, e.Level, e.Title, e.Kind, e.StartPage, e.EndPage, e.Note,
                bySource[e.Order].SourceUniqueId, bySource[e.Order].SourceItemId
            })
        };
        var manifestsDrive = await graph.GetListDriveIdAsync(plan.SiteId, plan.ManifestsListId, ct);
        var (_, manifestItemId) = await graph.UploadAsync(manifestsDrive, $"{plan.BundleId}/v{version}.json",
            JsonSerializer.SerializeToUtf8Bytes(manifest, ManifestJson), ct);
        await SetColumnsAsync(plan.SiteId, webUrl, plan.ManifestsListId, manifestItemId,
            new Dictionary<string, object> { ["BundleId"] = plan.BundleId, ["BuildVersion"] = version }, ct);

        await graph.UpdateListItemFieldsAsync(plan.SiteId, plan.RegistryListId, plan.RegistryItemId,
            new { BundleStatus = "Built", LatestPdfUrl = file.WebUrl }, ct);
        await work.DeletePrefixAsync(input.WorkPrefix, ct);

        return new BuildResult(file.WebUrl, version, input.Assembled.PageCount, input.Members.Count,
            input.Assembled.Entries.Count(e => e.Kind == MemberKind.Placeholder), input.Assembled.Entries.Count(e => e.Level > 0));
    }

    /// <summary>
    /// Sets columns on the current version without adding one (SharePoint REST). If the Function lacks the
    /// SharePoint Sites.Selected role, falls back to Graph, which adds a second version per build.
    /// </summary>
    private async Task SetColumnsAsync(string siteId, Uri webUrl, string listId, int itemId, IReadOnlyDictionary<string, object> values, CancellationToken ct)
    {
        try
        {
            await sharePoint.UpdateFieldsWithoutVersionAsync(webUrl, listId, itemId,
                values.ToDictionary(v => v.Key, v => Convert.ToString(v.Value, System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty), ct);
        }
        catch (HttpRequestException ex) when (ex.StatusCode is System.Net.HttpStatusCode.Unauthorized or System.Net.HttpStatusCode.Forbidden)
        {
            logger.LogWarning("SharePoint REST denied ({Status}); setting columns through Graph (adds a version). Grant the SharePoint Sites.Selected role.", ex.StatusCode);
            await graph.UpdateListItemFieldsAsync(siteId, listId, itemId, values, ct);
        }
    }

    private static string SafeFileName(string title)
    {
        var invalid = new HashSet<char>("\"*:<>?/\\|#%".ToCharArray());
        var cleaned = new string(title.Select(c => invalid.Contains(c) ? '_' : c).ToArray()).Trim().TrimEnd('.');
        return string.IsNullOrEmpty(cleaned) ? "Bundle" : cleaned;
    }
}

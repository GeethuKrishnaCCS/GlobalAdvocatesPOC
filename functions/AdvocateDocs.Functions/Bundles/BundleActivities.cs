using System.Text.Json;
using AdvocateDocs.Functions.Graph;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using UglyToad.PdfPig;

namespace AdvocateDocs.Functions.Bundles;

/// <summary>The bundle build's steps. Each is idempotent so Durable Functions can retry it.</summary>
public sealed class BundleActivities(GraphSiteClient graph, BundleWorkStore work, ILogger<BundleActivities> logger)
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

        return new BundlePlan(siteId, ListId(request.Lists.Documents), ListId(request.Lists.Bundles), registryId, ListId(request.Lists.Manifests),
            int.Parse(header["id"]!.GetValue<string>()), request.BundleId, header["fields"]!["Title"]!.GetValue<string>(), members);
    }

    /// <summary>Turns one member into PDF (as-is, Graph conversion, text rendering or a placeholder) and parks it in blob storage.</summary>
    [Function(nameof(ConvertMember))]
    public async Task<ConvertedMember> ConvertMember([ActivityTrigger] ConvertInput input, CancellationToken ct)
    {
        var m = input.Member;
        var blobName = $"{input.WorkPrefix}/members/{m.Order:D5}.pdf";
        var (pdf, kind, note) = await ToPdfAsync(input, ct);
        await work.PutAsync(blobName, pdf, ct);
        return new ConvertedMember(m.Order, m.Title, kind, blobName, m.SourceUniqueId, m.SourceItemId, note);
    }

    /// <summary>Builds the bundle PDF (index, stamps, bookmarks) from the parked members.</summary>
    [Function(nameof(AssembleBundle))]
    public async Task<AssembleOutput> AssembleBundle([ActivityTrigger] AssembleInput input, CancellationToken ct)
    {
        var members = new List<MemberPdf>(input.Members.Count);
        foreach (var m in input.Members)
        {
            members.Add(new MemberPdf(m.Order, m.Title, m.Kind, await work.GetAsync(m.BlobName, ct), m.Note));
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
        await graph.UpdateListItemFieldsAsync(plan.SiteId, plan.BundlesListId, fileItemId, new
        {
            BundleId = plan.BundleId, ManifestVersion = version, BuildStatus = "Ready", PageCount = input.Assembled.PageCount, VolumeNo = 1
        }, ct);

        var bySource = input.Members.ToDictionary(m => m.Order);
        var manifest = new
        {
            plan.BundleId, plan.Title, Version = version, BuiltAtUtc = DateTime.UtcNow, BuiltBy = input.Build.RequestedBy,
            input.Assembled.PageCount, input.Assembled.IndexPageCount, PdfUrl = file.WebUrl,
            Entries = input.Assembled.Entries.Select(e => new
            {
                e.Order, e.Title, e.Kind, e.StartPage, e.EndPage, e.Note,
                bySource[e.Order].SourceUniqueId, bySource[e.Order].SourceItemId
            })
        };
        var manifestsDrive = await graph.GetListDriveIdAsync(plan.SiteId, plan.ManifestsListId, ct);
        var (_, manifestItemId) = await graph.UploadAsync(manifestsDrive, $"{plan.BundleId}/v{version}.json",
            JsonSerializer.SerializeToUtf8Bytes(manifest, ManifestJson), ct);
        await graph.UpdateListItemFieldsAsync(plan.SiteId, plan.ManifestsListId, manifestItemId, new { BundleId = plan.BundleId, BuildVersion = version }, ct);

        await graph.UpdateListItemFieldsAsync(plan.SiteId, plan.RegistryListId, plan.RegistryItemId,
            new { BundleStatus = "Built", LatestPdfUrl = file.WebUrl }, ct);
        await work.DeletePrefixAsync(input.WorkPrefix, ct);

        return new BuildResult(file.WebUrl, version, input.Assembled.PageCount, input.Members.Count,
            input.Members.Count(m => m.Kind == MemberKind.Placeholder));
    }

    private async Task<(byte[] Pdf, string Kind, string? Note)> ToPdfAsync(ConvertInput input, CancellationToken ct)
    {
        var m = input.Member;
        if (m.SourceItemId <= 0)
        {
            return Placeholder(m, "No source reference stored for this member; re-save the bundle.");
        }
        try
        {
            var file = await graph.GetListItemDriveItemAsync(input.SiteId, input.DocumentsListId, m.SourceItemId, ct);
            var ext = Path.GetExtension(file.Name).TrimStart('.').ToLowerInvariant();
            byte[]? pdf;
            string kind;
            switch (ext)
            {
                case "pdf":
                    (pdf, kind) = (await graph.DownloadAsync(file.DriveId, file.ItemId, false, ct), MemberKind.Pdf);
                    break;
                case "txt":
                    var text = System.Text.Encoding.UTF8.GetString((await graph.DownloadAsync(file.DriveId, file.ItemId, false, ct))!);
                    (pdf, kind) = (TextPdfRenderer.RenderText(file.Name, text), MemberKind.Text);
                    break;
                case "eml":
                    (pdf, kind) = (TextPdfRenderer.RenderEmail(file.Name, (await graph.DownloadAsync(file.DriveId, file.ItemId, false, ct))!), MemberKind.Text);
                    break;
                default:
                    (pdf, kind) = (await graph.DownloadAsync(file.DriveId, file.ItemId, true, ct), MemberKind.Converted);
                    break;
            }
            if (pdf is null)
            {
                return Placeholder(m, $"'.{ext}' files can't be converted to PDF.", file.WebUrl);
            }
            if (!OpensAsPdf(pdf))
            {
                return Placeholder(m, "The file is not a readable PDF (it may be encrypted or damaged).", file.WebUrl);
            }
            return (pdf, kind, null);
        }
        catch (GraphRequestException ex) when (ex.Status is System.Net.HttpStatusCode.NotFound)
        {
            logger.LogWarning("Member {Order} source item {Item} not found", m.Order, m.SourceItemId);
            return Placeholder(m, "The source document no longer exists.");
        }
    }

    private static (byte[] Pdf, string Kind, string? Note) Placeholder(PlannedMember m, string reason, string? link = null)
        => (TextPdfRenderer.RenderPlaceholder(m.Title, reason, link), MemberKind.Placeholder, reason);

    private static bool OpensAsPdf(byte[] pdf)
    {
        try
        {
            using var doc = PdfDocument.Open(pdf);
            return doc.NumberOfPages > 0;
        }
        catch (Exception)
        {
            return false;
        }
    }

    private static string SafeFileName(string title)
    {
        var invalid = new HashSet<char>("\"*:<>?/\\|#%".ToCharArray());
        var cleaned = new string(title.Select(c => invalid.Contains(c) ? '_' : c).ToArray()).Trim().TrimEnd('.');
        return string.IsNullOrEmpty(cleaned) ? "Bundle" : cleaned;
    }
}

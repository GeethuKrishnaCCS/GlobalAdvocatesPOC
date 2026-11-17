using AdvocateDocs.Functions.Bundles;
using AdvocateDocs.Functions.Graph;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Options;

namespace AdvocateDocs.Functions.Functions;

/// <summary>What the SPFx Print command posts: documents library items to print, in order.</summary>
public sealed record PrintRequest(string WebUrl, string DocumentsList, string? ConversionList, IReadOnlyList<int> ItemIds);

/// <summary>Print: selected documents of any type as one A4 PDF (email attachments included), for the browser's print dialog.</summary>
public sealed class PrintFunctions(GraphSiteClient graph, MemberConverter converter, IOptions<SharePointOptions> sharePoint)
{
    /// <summary>Synchronous request: keeps it well inside the 230 s HTTP limit for typical documents.</summary>
    public const int MaxItems = 25;

    /// <summary>POST /api/print — body: <see cref="PrintRequest"/>. Returns application/pdf.</summary>
    [Function("Print")]
    public async Task<IActionResult> Print(
        [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "print")] HttpRequest request, FunctionContext context)
    {
        var ct = context.CancellationToken;
        var body = await request.ReadFromJsonAsync<PrintRequest>(ct);
        if (body?.ItemIds is not { Count: > 0 and <= MaxItems } || body.ItemIds.Any(id => id <= 0))
        {
            return new BadRequestObjectResult(new { error = $"Select 1 to {MaxItems} documents." });
        }
        var problem = BundleBuildFunctions.ValidateSite(body.WebUrl, sharePoint.Value.Host)
            ?? BundleBuildFunctions.ValidateListNames(body.ConversionList is null ? [body.DocumentsList] : [body.DocumentsList, body.ConversionList]);
        if (problem is not null)
        {
            return new BadRequestObjectResult(new { error = problem });
        }

        var siteId = await graph.GetSiteIdAsync(new Uri(body.WebUrl), ct);
        var lists = await graph.GetListIdsByNameAsync(siteId, ct);
        if (!lists.TryGetValue(body.DocumentsList, out var documentsId))
        {
            return new NotFoundObjectResult(new { error = $"List '{body.DocumentsList}' not found." });
        }
        var conversionId = body.ConversionList is not null && lists.TryGetValue(body.ConversionList, out var id) ? id : null;
        var workPrefix = $"print-{Guid.NewGuid():N}";

        var pdfs = new List<byte[]>();
        for (var i = 0; i < body.ItemIds.Count; i++)
        {
            var member = new PlannedMember(i + 1, body.ItemIds[i], string.Empty, string.Empty);
            var c = await converter.ConvertAsync(new ConvertInput(workPrefix, siteId, documentsId, member, conversionId), ct);
            pdfs.Add(c.Pdf);
            pdfs.AddRange(c.Attachments.Select(a => a.Pdf));
        }
        return new FileContentResult(PdfBundleAssembler.Concatenate(pdfs), "application/pdf");
    }
}

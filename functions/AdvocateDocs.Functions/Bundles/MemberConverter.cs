using System.Globalization;
using AdvocateDocs.Functions.Graph;
using Microsoft.Extensions.Logging;
using UglyToad.PdfPig;

namespace AdvocateDocs.Functions.Bundles;

/// <summary>One document as A4 PDF pages, plus its email attachments (if any), in order.</summary>
public sealed record MemberConversion(byte[] Pdf, string Kind, string? Note, IReadOnlyList<AttachmentPdf> Attachments);

/// <summary>
/// Turns a matter document into A4 PDF pages: as-is, Graph conversion, text or image rendering, or a placeholder
/// page saying why it could not be included. Shared by bundle builds and Print.
/// </summary>
public sealed class MemberConverter(GraphSiteClient graph, EmailBundleConverter email, ILogger<MemberConverter> logger)
{
    public async Task<MemberConversion> ConvertAsync(ConvertInput input, CancellationToken ct)
    {
        var c = await ToPdfAsync(input, ct);
        return c with
        {
            Pdf = PageNormalizer.ToA4(c.Pdf),
            Attachments = c.Attachments.Select(a => a with { Pdf = PageNormalizer.ToA4(a.Pdf) }).ToList()
        };
    }

    private async Task<MemberConversion> ToPdfAsync(ConvertInput input, CancellationToken ct)
    {
        var m = input.Member;
        if (m.SourceItemId <= 0)
        {
            return Placeholder(m.Title, "No source reference stored for this member; re-save the bundle.");
        }
        try
        {
            var file = await graph.GetListItemDriveItemAsync(input.SiteId, input.DocumentsListId, m.SourceItemId, ct);
            var title = string.IsNullOrEmpty(m.Title) ? file.Name : m.Title;
            var ext = Path.GetExtension(file.Name).TrimStart('.').ToLowerInvariant();
            byte[]? pdf;
            string kind;
            switch (ext)
            {
                case "pdf":
                    (pdf, kind) = (await graph.DownloadAsync(file.DriveId, file.ItemId, false, ct), MemberKind.Pdf);
                    break;
                case "txt" or "csv" or "log" or "xml" or "json":
                    var text = System.Text.Encoding.UTF8.GetString((await graph.DownloadAsync(file.DriveId, file.ItemId, false, ct))!);
                    (pdf, kind) = (TextPdfRenderer.RenderText(file.Name, text), MemberKind.Text);
                    break;
                case "eml":
                    // Body like SharePoint's preview, then each real attachment as its own index entry ("4.1", "4.2"...).
                    var eml = (await graph.DownloadAsync(file.DriveId, file.ItemId, false, ct))!;
                    var target = new ConversionTarget(input.SiteId, input.ConversionListId, input.WorkPrefix);
                    var (body, bodyKind) = await email.RenderBodyAsync(file, eml, target, ct);
                    var attachments = await email.ConvertAttachmentsAsync(eml, m.Order.ToString(CultureInfo.InvariantCulture), 1, target, file.WebUrl, ct);
                    return new MemberConversion(body, bodyKind, null, attachments);
                case "png" or "jpg" or "jpeg":
                    (pdf, kind) = (ImagePdfRenderer.Render((await graph.DownloadAsync(file.DriveId, file.ItemId, false, ct))!), MemberKind.Image);
                    break;
                case "vsdx" or "vsd":
                    // No Visio-to-PDF conversion in M365: use SharePoint's preview image of the first page, when it has one.
                    var preview = await graph.DownloadPreviewImageAsync(file.DriveId, file.ItemId, ct);
                    if (preview is null)
                    {
                        return Placeholder(title, "Microsoft 365 can't convert Visio drawings to PDF. Open the original, or save it as PDF in Visio and add that file.", file.WebUrl);
                    }
                    (pdf, kind) = (ImagePdfRenderer.Render(preview, "Preview image of the first page only - open the original for full detail."), MemberKind.Preview);
                    break;
                default:
                    (pdf, kind) = (await graph.DownloadAsync(file.DriveId, file.ItemId, true, ct), MemberKind.Converted);
                    break;
            }
            if (pdf is null)
            {
                return Placeholder(title, $"'.{ext}' files can't be converted to PDF.", file.WebUrl);
            }
            if (!OpensAsPdf(pdf))
            {
                return Placeholder(title, "The file is not a readable PDF (it may be encrypted or damaged).", file.WebUrl);
            }
            return new MemberConversion(pdf, kind, null, []);
        }
        catch (GraphRequestException ex) when (ex.Status is System.Net.HttpStatusCode.NotFound)
        {
            logger.LogWarning("Member {Order} source item {Item} not found", m.Order, m.SourceItemId);
            return Placeholder(m.Title, "The source document no longer exists.");
        }
        catch (GraphRequestException ex)
        {
            // One unconvertible file (password-protected, damaged, too large) must not fail the whole bundle;
            // throttling is already retried inside GraphSiteClient.
            logger.LogWarning("Member {Order} source item {Item} not converted: {Error}", m.Order, m.SourceItemId, ex.Message);
            return Placeholder(m.Title, $"Microsoft 365 could not convert the file ({ex.Message}).");
        }
    }

    private static MemberConversion Placeholder(string title, string reason, string? link = null)
        => new(TextPdfRenderer.RenderPlaceholder(title, reason, link), MemberKind.Placeholder, reason, []);

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
}

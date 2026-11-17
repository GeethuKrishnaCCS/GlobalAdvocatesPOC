using System.Globalization;
using AdvocateDocs.Functions.Graph;
using Microsoft.Extensions.Logging;
using MimeKit;
using UglyToad.PdfPig;

namespace AdvocateDocs.Functions.Bundles;

/// <summary>An email attachment rendered to PDF, ready to follow its email in the bundle.</summary>
public sealed record AttachmentPdf(string Label, int Level, string Title, string Kind, byte[] Pdf, string? Note);

/// <summary>Where Graph can convert loose bytes: a (hidden) library on the matter site, used for short-lived copies.</summary>
public sealed record ConversionTarget(string SiteId, string? ListId, string WorkPrefix);

/// <summary>
/// Emails in bundles: the body rendered like SharePoint's own email preview (Graph PDF conversion: headers, Cc,
/// HTML layout, inline images and signature logos), then each real attachment as its own index entry.
/// Attached emails are expanded the same way, down to <see cref="MaxDepth"/> levels.
/// </summary>
public sealed class EmailBundleConverter(GraphSiteClient graph, ILogger<EmailBundleConverter> logger)
{
    /// <summary>Attachment levels expanded below the bundle member (4.1 = level 1, 4.1.1 = level 2).</summary>
    public const int MaxDepth = 2;

    private static readonly HashSet<string> TextTypes = new(StringComparer.OrdinalIgnoreCase) { "txt", "csv", "log", "xml", "json" };
    private static readonly HashSet<string> ImageTypes = new(StringComparer.OrdinalIgnoreCase) { "png", "jpg", "jpeg" };

    /// <summary>
    /// The email's own pages: Graph's conversion (of a copy with inline images embedded, when it has any; else of the
    /// stored file), else our text rendering.
    /// </summary>
    public async Task<(byte[] Pdf, string Kind)> RenderBodyAsync(DriveItemRef file, byte[] eml, ConversionTarget target, CancellationToken ct)
    {
        var converted = await ConvertEmailAsync(file.Name, eml, () => graph.DownloadAsync(file.DriveId, file.ItemId, true, ct), target, ct);
        return converted is not null ? (converted, MemberKind.Converted) : (TextPdfRenderer.RenderEmail(file.Name, eml), MemberKind.Text);
    }

    /// <summary>
    /// Graph's conversion of an email with its inline images embedded; images too large to embed follow as pages.
    /// Without inline images (or a conversion library) the stored file is converted as it is. Null when Graph can't.
    /// </summary>
    private async Task<byte[]?> ConvertEmailAsync(string name, byte[] eml, Func<Task<byte[]?>>? convertStored, ConversionTarget target, CancellationToken ct)
    {
        (byte[] Eml, IReadOnlyList<EmailAttachment> Oversized)? embedded;
        using (var stream = new MemoryStream(eml))
        {
            embedded = EmailAttachments.WithInlineImagesEmbedded(await MimeMessage.LoadAsync(stream, ct));
        }
        var body = embedded is { } e && target.ListId is not null
            ? await ConvertWithGraphAsync(new EmailAttachment(name, e.Eml, IsMessage: true), target, ct)
            : convertStored is not null
                ? await TryConvertAsync(convertStored, name)
                : await ConvertWithGraphAsync(new EmailAttachment(name, eml, IsMessage: true), target, ct);
        var pages = (embedded?.Oversized ?? [])
            .Select(o => ImagePdfRenderer.Render(o.Content, $"Image from the email body: {o.FileName}"))
            .OfType<byte[]>()
            .ToList();
        return body is null || pages.Count == 0 ? body : PdfBundleAssembler.Concatenate([body, .. pages]);
    }

    /// <summary>Attachments of <paramref name="eml"/> as PDFs, labelled under <paramref name="parentLabel"/>, attached emails expanded.</summary>
    public async Task<IReadOnlyList<AttachmentPdf>> ConvertAttachmentsAsync(byte[] eml, string parentLabel, int level,
        ConversionTarget target, string? parentLink, CancellationToken ct)
    {
        if (level > MaxDepth)
        {
            return [];
        }
        MimeMessage message;
        using (var stream = new MemoryStream(eml))
        {
            message = await MimeMessage.LoadAsync(stream, ct);
        }
        var result = new List<AttachmentPdf>();
        var attachments = EmailAttachments.Extract(message);
        for (var i = 0; i < attachments.Count; i++)
        {
            var a = attachments[i];
            var label = $"{parentLabel}.{(i + 1).ToString(CultureInfo.InvariantCulture)}";
            var (pdf, kind, note) = await ConvertAsync(a, target, ct);
            if (kind == MemberKind.Placeholder)
            {
                pdf = TextPdfRenderer.RenderPlaceholder(a.FileName, note!, parentLink);
            }
            result.Add(new AttachmentPdf(label, level, a.FileName, kind, pdf, note));
            if (a.IsMessage)
            {
                result.AddRange(await ConvertAttachmentsAsync(a.Content, label, level + 1, target, parentLink, ct));
            }
        }
        return result;
    }

    private async Task<(byte[] Pdf, string Kind, string? Note)> ConvertAsync(EmailAttachment a, ConversionTarget target, CancellationToken ct)
    {
        var ext = Path.GetExtension(a.FileName).TrimStart('.');
        if (a.IsMessage)
        {
            var converted = await ConvertEmailAsync(a.FileName, a.Content, null, target, ct);
            return converted is not null ? (converted, MemberKind.Converted, null) : (TextPdfRenderer.RenderEmail(a.FileName, a.Content), MemberKind.Text, null);
        }
        if (ext.Equals("pdf", StringComparison.OrdinalIgnoreCase))
        {
            return OpensAsPdf(a.Content) ? (a.Content, MemberKind.Pdf, null) : Placeholder("The attached PDF is not readable (it may be encrypted or damaged).");
        }
        if (TextTypes.Contains(ext))
        {
            return (TextPdfRenderer.RenderText(a.FileName, System.Text.Encoding.UTF8.GetString(a.Content)), MemberKind.Text, null);
        }
        if (ImageTypes.Contains(ext))
        {
            var image = ImagePdfRenderer.Render(a.Content);
            return image is not null ? (image, MemberKind.Image, null) : Placeholder("The attached image could not be read.");
        }
        if (target.ListId is null)
        {
            return Placeholder("No conversion library is configured for this matter site.");
        }
        var pdf = await ConvertWithGraphAsync(a, target, ct);
        return pdf is not null ? (pdf, MemberKind.Converted, null) : Placeholder($"'.{ext}' attachments can't be converted to PDF.");
    }

    /// <summary>Graph converts only stored files: park a copy in the conversion library, convert, delete the copy.</summary>
    private async Task<byte[]?> ConvertWithGraphAsync(EmailAttachment a, ConversionTarget target, CancellationToken ct)
    {
        if (target.ListId is null)
        {
            return null;
        }
        var drive = await graph.GetListDriveIdAsync(target.SiteId, target.ListId, ct);
        var ext = a.IsMessage ? ".eml" : Path.GetExtension(a.FileName);
        var (copy, _) = await graph.UploadAsync(drive, $"{target.WorkPrefix}/{Guid.NewGuid():N}{ext}", a.Content, ct);
        try
        {
            return await TryConvertAsync(() => graph.DownloadAsync(copy.DriveId, copy.ItemId, true, ct), a.FileName);
        }
        finally
        {
            try
            {
                await graph.DeleteAsync(copy.DriveId, copy.ItemId, ct);
            }
            catch (GraphRequestException ex)
            {
                // A leftover copy sits only in the hidden library; never fail the build for it.
                logger.LogWarning("Temporary copy {Item} not deleted: {Error}", copy.ItemId, ex.Message);
            }
        }
    }

    /// <summary>A readable PDF from a Graph conversion, or null (unsupported type, conversion error, unreadable output).</summary>
    private async Task<byte[]?> TryConvertAsync(Func<Task<byte[]?>> convert, string name)
    {
        try
        {
            var pdf = await convert();
            return pdf is not null && OpensAsPdf(pdf) ? pdf : null;
        }
        catch (GraphRequestException ex)
        {
            logger.LogWarning("Graph could not convert {Name}: {Error}", name, ex.Message);
            return null;
        }
    }

    private static (byte[] Pdf, string Kind, string? Note) Placeholder(string reason) => ([], MemberKind.Placeholder, reason);

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

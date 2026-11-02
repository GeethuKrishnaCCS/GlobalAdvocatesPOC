using System.Text.RegularExpressions;
using MimeKit;

namespace AdvocateDocs.Functions.Bundles;

/// <summary>An email attachment to add to a bundle after its email.</summary>
/// <param name="IsMessage">An attached email (message/rfc822); its own attachments can be expanded in turn.</param>
public sealed record EmailAttachment(string FileName, byte[] Content, bool IsMessage);

/// <summary>
/// Finds an email's real attachments: files and attached emails, in message order. Images the HTML body shows
/// through <c>cid:</c> (signature logos, pasted pictures) are part of the body, not exhibits, and are skipped —
/// whatever their Content-Disposition, since Outlook often marks signature logos as attachments.
/// </summary>
public static partial class EmailAttachments
{
    /// <summary>More than this per email is unusual; the rest are left out (the email body still lists them).</summary>
    public const int MaxPerEmail = 50;

    /// <summary>
    /// shortcut: Graph's conversion fails (412) on very large data URIs (2.7 MB PNG failed, 300 KB JPEG fine; exact
    /// cap unknown). Bigger inline images go on their own page after the email instead; re-encode them (e.g. SkiaSharp)
    /// if the firm wants them in place.
    /// </summary>
    public const int MaxInlineImageBytes = 512 * 1024;

    [GeneratedRegex(@"cid:([^""'\s>)]+)", RegexOptions.IgnoreCase)]
    private static partial Regex CidReference();

    [GeneratedRegex(@"<img\b[^>]*?cid:([^""'\s>)]+)[^>]*>", RegexOptions.IgnoreCase)]
    private static partial Regex CidImageTag();

    public static IReadOnlyList<EmailAttachment> Extract(MimeMessage message)
    {
        var referenced = ReferencedContentIds(message.HtmlBody);
        var result = new List<EmailAttachment>();
        foreach (var entity in message.BodyParts)
        {
            if (result.Count == MaxPerEmail)
            {
                break;
            }
            if (entity.ContentId is { } cid && referenced.Contains(cid))
            {
                continue;
            }
            switch (entity)
            {
                case MessagePart { Message: { } inner } part:
                    result.Add(new EmailAttachment(MessageFileName(part, inner), ToBytes(inner), IsMessage: true));
                    break;
                case MimePart { Content: not null } part when IsFile(part):
                    using (var stream = new MemoryStream())
                    {
                        part.Content!.DecodeTo(stream);
                        result.Add(new EmailAttachment(part.FileName ?? $"attachment-{result.Count + 1}", stream.ToArray(), IsMessage: false));
                    }
                    break;
            }
        }
        return result;
    }

    /// <summary>
    /// The email with its <c>cid:</c> images written into the HTML as data URIs (Graph's PDF conversion can't resolve
    /// cid references, but renders data URIs). Images over <see cref="MaxInlineImageBytes"/> are replaced by a note
    /// and returned in <c>Oversized</c>, to follow the email as pages. Null when the HTML references no inline parts.
    /// </summary>
    public static (byte[] Eml, IReadOnlyList<EmailAttachment> Oversized)? WithInlineImagesEmbedded(MimeMessage message)
    {
        var html = message.BodyParts.OfType<TextPart>().FirstOrDefault(t => t.IsHtml && !t.IsAttachment);
        var images = message.BodyParts.OfType<MimePart>()
            .Where(p => p.ContentId is not null && p.Content is not null)
            .GroupBy(p => p.ContentId!, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => (Part: g.First(), Bytes: Decode(g.First())), StringComparer.OrdinalIgnoreCase);
        if (html is null || images.Count == 0)
        {
            return null;
        }
        var replaced = 0;
        var oversized = new List<EmailAttachment>();
        var text = CidImageTag().Replace(html.Text, m =>
        {
            if (!images.TryGetValue(Uri.UnescapeDataString(m.Groups[1].Value), out var image) || image.Bytes.Length <= MaxInlineImageBytes)
            {
                return m.Value;
            }
            replaced++;
            var name = image.Part.FileName ?? $"inline-image-{oversized.Count + 1}";
            oversized.Add(new EmailAttachment(name, image.Bytes, IsMessage: false));
            return $"<p><i>[Image {System.Net.WebUtility.HtmlEncode(name)}: shown on the following page]</i></p>";
        });
        var used = new HashSet<MimeEntity>();
        html.Text = CidReference().Replace(text, m =>
        {
            if (!images.TryGetValue(Uri.UnescapeDataString(m.Groups[1].Value), out var image))
            {
                return m.Value;
            }
            replaced++;
            used.Add(image.Part);
            return $"data:{image.Part.ContentType.MimeType};base64,{Convert.ToBase64String(image.Bytes)}";
        });
        used.UnionWith(oversized.Select(o => images.Values.First(i => i.Bytes == o.Content).Part));
        // Embedded parts would otherwise reappear in the converted header's "Attachments:" line.
        RemoveParts(message.Body, used);
        // Graph's converter shows quoted-printable soft breaks literally ("follo=wing") in re-encoded parts; base64 doesn't break text.
        html.ContentTransferEncoding = ContentEncoding.Base64;
        return replaced == 0 ? null : (ToBytes(message), oversized);
    }

    private static void RemoveParts(MimeEntity? entity, HashSet<MimeEntity> parts)
    {
        if (entity is not Multipart multipart)
        {
            return;
        }
        for (var i = multipart.Count - 1; i >= 0; i--)
        {
            if (parts.Contains(multipart[i]))
            {
                multipart.RemoveAt(i);
            }
            else
            {
                RemoveParts(multipart[i], parts);
            }
        }
    }

    private static byte[] Decode(MimePart part)
    {
        using var stream = new MemoryStream();
        part.Content!.DecodeTo(stream);
        return stream.ToArray();
    }

    /// <summary>Names of attachments, for the "Attachments:" line of a rendered email.</summary>
    public static IReadOnlyList<string> Names(MimeMessage message) => Extract(message).Select(a => a.FileName).ToList();

    /// <summary>A file part, not the text/HTML body: marked as an attachment or carrying a file name.</summary>
    private static bool IsFile(MimePart part) => part.IsAttachment || !string.IsNullOrEmpty(part.FileName);

    private static HashSet<string> ReferencedContentIds(string? html)
    {
        var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (!string.IsNullOrEmpty(html))
        {
            foreach (Match m in CidReference().Matches(html))
            {
                ids.Add(Uri.UnescapeDataString(m.Groups[1].Value));
            }
        }
        return ids;
    }

    private static string MessageFileName(MessagePart part, MimeMessage inner)
    {
        var name = part.ContentDisposition?.FileName;
        if (!string.IsNullOrEmpty(name))
        {
            return name.EndsWith(".eml", StringComparison.OrdinalIgnoreCase) ? name : $"{name}.eml";
        }
        var subject = string.IsNullOrWhiteSpace(inner.Subject) ? "Attached email" : inner.Subject;
        return $"{string.Concat(subject.Select(c => Path.GetInvalidFileNameChars().Contains(c) ? '_' : c))}.eml";
    }

    private static byte[] ToBytes(MimeMessage message)
    {
        using var stream = new MemoryStream();
        message.WriteTo(stream);
        return stream.ToArray();
    }
}

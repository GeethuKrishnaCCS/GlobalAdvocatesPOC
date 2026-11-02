using MimeKit;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Core;
using UglyToad.PdfPig.Fonts.Standard14Fonts;
using UglyToad.PdfPig.Writer;

namespace AdvocateDocs.Functions.Bundles;

/// <summary>Renders formats Graph can't convert (txt, eml) and placeholder pages as simple A4 PDFs.</summary>
public static class TextPdfRenderer
{
    private const double Margin = 48;
    private const double TitleSize = 13;
    private const double BodySize = 9;
    private const double LineHeight = 11.5;
    private const int CharsPerLine = 95;

    /// <summary>Plain text, monospace, wrapped; the title heads the first page.</summary>
    public static byte[] RenderText(string title, string text)
        => Render(title, text.Replace("\r\n", "\n").Split('\n'));

    /// <summary>An email: key headers, then the text body (HTML body reduced to text when there is no text part).</summary>
    public static byte[] RenderEmail(string title, byte[] eml)
    {
        using var stream = new MemoryStream(eml);
        var message = MimeMessage.Load(stream);
        var body = message.TextBody ?? StripTags(message.HtmlBody ?? string.Empty);
        var lines = new List<string>
        {
            $"From:    {message.From}",
            $"To:      {message.To}",
        };
        if (message.Cc.Count > 0)
        {
            lines.Add($"Cc:      {message.Cc}");
        }
        lines.Add($"Date:    {message.Date:yyyy-MM-dd HH:mm zzz}");
        lines.Add($"Subject: {message.Subject}");
        var attachments = EmailAttachments.Names(message);
        if (attachments.Count > 0)
        {
            lines.Add($"Attachments: {string.Join(", ", attachments)}");
        }
        lines.Add(new string('-', CharsPerLine));
        lines.AddRange(body.Replace("\r\n", "\n").Split('\n'));
        return Render(title, lines);
    }

    /// <summary>One page saying the document could not be included and why — it keeps its place in the index.</summary>
    public static byte[] RenderPlaceholder(string title, string reason, string? link)
    {
        var lines = new List<string> { "This document could not be included in the bundle.", string.Empty, $"Reason: {reason}" };
        if (!string.IsNullOrEmpty(link))
        {
            lines.Add(string.Empty);
            lines.Add($"Open the original: {link}");
        }
        return Render(title, lines);
    }

    private static byte[] Render(string title, IEnumerable<string> lines)
    {
        using var builder = new PdfDocumentBuilder();
        var bold = builder.AddStandard14Font(Standard14Font.HelveticaBold);
        var mono = builder.AddStandard14Font(Standard14Font.Courier);

        PdfPageBuilder? page = null;
        double y = 0;
        void NewPage()
        {
            page = builder.AddPage(PageSize.A4);
            y = page.PageSize.Height - Margin;
        }

        NewPage();
        page!.AddText(PdfText.Truncate(PdfText.ToWinAnsi(title), 80), TitleSize, new PdfPoint(Margin, y), bold);
        y -= TitleSize * 2;
        foreach (var line in lines.SelectMany(l => PdfText.Wrap(PdfText.ToWinAnsi(l), CharsPerLine)))
        {
            if (y < Margin)
            {
                NewPage();
            }
            page!.AddText(line, BodySize, new PdfPoint(Margin, y), mono);
            y -= LineHeight;
        }
        return builder.Build();
    }

    private static string StripTags(string html)
        => System.Text.RegularExpressions.Regex.Replace(html, "<[^>]+>", " ").Replace("&nbsp;", " ");
}

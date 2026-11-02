using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace AdvocateDocs.Functions.Bundles;

/// <summary>
/// Puts every page on an A4 sheet: portrait pages on A4 portrait, landscape pages on A4 landscape, content scaled to
/// fit and centred. Pages already A4 are kept as they are.
/// </summary>
public static class PageNormalizer
{
    public const double A4Width = 595.28;
    public const double A4Height = 841.89;
    private const double Tolerance = 2;

    /// <summary>The PDF with A4 pages; the input unchanged when it already is, or when PDFsharp can't read it.</summary>
    public static byte[] ToA4(byte[] pdf)
    {
        try
        {
            using var source = PdfReader.Open(new MemoryStream(pdf), PdfDocumentOpenMode.Import);
            if (source.Pages.Cast<PdfPage>().All(IsA4))
            {
                return pdf;
            }
            using var form = XPdfForm.FromStream(new MemoryStream(pdf));
            using var output = new PdfDocument();
            for (var i = 0; i < source.PageCount; i++)
            {
                var page = source.Pages[i];
                if (IsA4(page))
                {
                    output.AddPage(page);
                    continue;
                }
                form.PageNumber = i + 1;
                var landscape = form.PointWidth > form.PointHeight;
                var (width, height) = landscape ? (A4Height, A4Width) : (A4Width, A4Height);
                var target = output.AddPage();
                target.Width = XUnit.FromPoint(width);
                target.Height = XUnit.FromPoint(height);
                var scale = Math.Min(width / form.PointWidth, height / form.PointHeight);
                var (w, h) = (form.PointWidth * scale, form.PointHeight * scale);
                using var gfx = XGraphics.FromPdfPage(target);
                gfx.DrawImage(form, (width - w) / 2, (height - h) / 2, w, h);
            }
            using var stream = new MemoryStream();
            output.Save(stream);
            return stream.ToArray();
        }
        catch (Exception ex) when (ex is PdfReaderException or InvalidOperationException or NotImplementedException)
        {
            // shortcut: unreadable-by-PDFsharp PDFs keep their own page size; upgrade if real matters hit this often.
            return pdf;
        }
    }

    private static bool IsA4(PdfPage page)
    {
        var (w, h) = (page.Width.Point, page.Height.Point);
        if (page.Rotate % 180 != 0)
        {
            (w, h) = (h, w);
        }
        return Near(Math.Min(w, h), A4Width) && Near(Math.Max(w, h), A4Height);
    }

    private static bool Near(double a, double b) => Math.Abs(a - b) <= Tolerance;
}

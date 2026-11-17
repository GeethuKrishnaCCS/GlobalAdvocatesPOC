using System.Buffers.Binary;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Core;
using UglyToad.PdfPig.Fonts.Standard14Fonts;
using UglyToad.PdfPig.Writer;

namespace AdvocateDocs.Functions.Bundles;

/// <summary>Puts a PNG or JPEG on one A4 page, scaled to fit (never enlarged), under an optional caption.</summary>
public static class ImagePdfRenderer
{
    private const double Margin = 36;
    private const double CaptionSize = 9;

    /// <summary>One-page PDF of the image; null when the bytes are not a PNG or JPEG.</summary>
    public static byte[]? Render(byte[] image, string? caption = null)
    {
        if (Dimensions(image) is not var (width, height, isPng))
        {
            return null;
        }
        using var builder = new PdfDocumentBuilder();
        var page = builder.AddPage(PageSize.A4);
        var top = page.PageSize.Height - Margin;
        if (!string.IsNullOrEmpty(caption))
        {
            var font = builder.AddStandard14Font(Standard14Font.HelveticaOblique);
            page.AddText(PdfText.Truncate(PdfText.ToWinAnsi(caption), 110), CaptionSize, new PdfPoint(Margin, top - CaptionSize), font);
            top -= CaptionSize * 2.5;
        }
        var boxWidth = page.PageSize.Width - 2 * Margin;
        var boxHeight = top - Margin;
        // Pixels as points at 72 dpi; shrink to fit the box, keep aspect ratio.
        var scale = Math.Min(1, Math.Min(boxWidth / width, boxHeight / height));
        var (w, h) = (width * scale, height * scale);
        var placement = new PdfRectangle(Margin, top - h, Margin + w, top);
        if (isPng)
        {
            page.AddPng(image, placement);
        }
        else
        {
            page.AddJpeg(image, placement);
        }
        return builder.Build();
    }

    private static (int Width, int Height, bool IsPng)? Dimensions(byte[] b)
    {
        // PNG: 8-byte signature, then IHDR with big-endian width/height at offsets 16/20.
        if (b.Length >= 24 && b[0] == 0x89 && b[1] == 'P' && b[2] == 'N' && b[3] == 'G')
        {
            return (BinaryPrimitives.ReadInt32BigEndian(b.AsSpan(16)), BinaryPrimitives.ReadInt32BigEndian(b.AsSpan(20)), true);
        }
        if (b.Length < 4 || b[0] != 0xFF || b[1] != 0xD8)
        {
            return null;
        }
        // JPEG: walk segments to the first start-of-frame (SOF0..SOF15 except DHT C4, JPG C8, DAC CC).
        var i = 2;
        while (i + 9 < b.Length)
        {
            if (b[i] != 0xFF)
            {
                return null;
            }
            var marker = b[i + 1];
            if (marker is >= 0xC0 and <= 0xCF and not 0xC4 and not 0xC8 and not 0xCC)
            {
                return (BinaryPrimitives.ReadUInt16BigEndian(b.AsSpan(i + 7)), BinaryPrimitives.ReadUInt16BigEndian(b.AsSpan(i + 5)), false);
            }
            i += 2 + BinaryPrimitives.ReadUInt16BigEndian(b.AsSpan(i + 2));
        }
        return null;
    }
}

using AdvocateDocs.Functions.Bundles;
using UglyToad.PdfPig;

namespace AdvocateDocs.Functions.Tests;

public class ImagePdfRendererTests
{
    /// <summary>Smallest valid PNG: 1x1 pixel, 8-bit RGB.</summary>
    private static readonly byte[] OnePixelPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGP4z8AAAAMBAQDJ/pLvAAAAAElFTkSuQmCC");

    [Fact]
    public void Render_Should_ProduceOnePagePdfWithImage_When_Png()
    {
        var pdf = ImagePdfRenderer.Render(OnePixelPng, "Caption");

        Assert.NotNull(pdf);
        using var doc = PdfDocument.Open(pdf);
        Assert.Equal(1, doc.NumberOfPages);
        Assert.Single(doc.GetPage(1).GetImages());
        Assert.Contains("Caption", doc.GetPage(1).Text);
    }

    [Fact]
    public void Render_Should_ReturnNull_When_NotPngOrJpeg()
    {
        Assert.Null(ImagePdfRenderer.Render("GIF89a not an image we embed"u8.ToArray()));
    }
}

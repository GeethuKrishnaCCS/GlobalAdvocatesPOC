using AdvocateDocs.Functions.Bundles;
using MimeKit;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Core;
using UglyToad.PdfPig.Fonts.Standard14Fonts;
using UglyToad.PdfPig.Writer;

namespace AdvocateDocs.Functions.Tests;

public class PageNormalizerTests
{
    private static byte[] PdfWithPages(params (double Width, double Height)[] sizes)
    {
        using var builder = new PdfDocumentBuilder();
        var font = builder.AddStandard14Font(Standard14Font.Helvetica);
        foreach (var (w, h) in sizes)
        {
            builder.AddPage(w, h).AddText("content", 12, new PdfPoint(20, h - 40), font);
        }
        return builder.Build();
    }

    [Fact]
    public void ToA4_Should_PutLetterOnA4PortraitAndWidePagesOnA4Landscape_When_SizesDiffer()
    {
        var pdf = PageNormalizer.ToA4(PdfWithPages((612, 792), (1224, 792), (595.28, 841.89)));

        using var doc = PdfDocument.Open(pdf);
        var sizes = Enumerable.Range(1, doc.NumberOfPages).Select(p => doc.GetPage(p)).Select(p => (Math.Round(p.Width), Math.Round(p.Height))).ToList();
        Assert.Equal([(595d, 842d), (842d, 595d), (595d, 842d)], sizes);
        Assert.Contains("content", doc.GetPage(1).Text);
    }

    [Fact]
    public void ToA4_Should_ReturnInputUnchanged_When_AllPagesAreA4()
    {
        var pdf = PdfWithPages((595.28, 841.89));

        Assert.Same(pdf, PageNormalizer.ToA4(pdf));
    }

    [Fact]
    public void WithInlineImagesEmbedded_Should_ReplaceCidWithDataUri_When_HtmlReferencesInlineImage()
    {
        var builder = new BodyBuilder { HtmlBody = "<p>Regards</p><img src=\"cid:logo@x\"/>" };
        var logo = builder.LinkedResources.Add("image001.png", [1, 2, 3], new ContentType("image", "png"));
        logo.ContentId = "logo@x";
        var message = new MimeMessage { Body = builder.ToMessageBody() };

        var embedded = EmailAttachments.WithInlineImagesEmbedded(message);

        Assert.NotNull(embedded);
        var html = MimeMessage.Load(new MemoryStream(embedded.Value.Eml)).HtmlBody;
        Assert.Contains("src=\"data:image/png;base64,AQID\"", html);
        Assert.DoesNotContain("cid:", html);
        Assert.Empty(embedded.Value.Oversized);
        Assert.DoesNotContain(MimeMessage.Load(new MemoryStream(embedded.Value.Eml)).BodyParts, p => p.ContentId == "logo@x");
    }

    [Fact]
    public void WithInlineImagesEmbedded_Should_MoveImageToOversized_When_ItExceedsTheEmbedLimit()
    {
        var builder = new BodyBuilder { HtmlBody = "<p>Notice</p><img src=\"cid:poster@x\" width=\"600\"/>" };
        var poster = builder.LinkedResources.Add("poster.png", new byte[EmailAttachments.MaxInlineImageBytes + 1], new ContentType("image", "png"));
        poster.ContentId = "poster@x";
        var message = new MimeMessage { Body = builder.ToMessageBody() };

        var embedded = EmailAttachments.WithInlineImagesEmbedded(message);

        Assert.NotNull(embedded);
        Assert.Equal("poster.png", Assert.Single(embedded.Value.Oversized).FileName);
        var html = MimeMessage.Load(new MemoryStream(embedded.Value.Eml)).HtmlBody;
        Assert.Contains("shown on the following page", html);
        Assert.DoesNotContain("<img", html);
    }
}

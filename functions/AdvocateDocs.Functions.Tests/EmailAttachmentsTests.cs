using System.Text;
using AdvocateDocs.Functions.Bundles;
using MimeKit;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Outline;

namespace AdvocateDocs.Functions.Tests;

public class EmailAttachmentsTests
{
    private static readonly byte[] Png = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");

    /// <summary>Outlook-style email: HTML body with a cid-referenced logo (marked as attachment), plus real attachments.</summary>
    private static MimeMessage EmailWithSignatureLogo()
    {
        var builder = new BodyBuilder();
        var logo = builder.Attachments.Add("image001.png", Png, new ContentType("image", "png"));
        logo.ContentId = "image001.png@01DB2F1A";
        builder.HtmlBody = "<p>Invoice attached.</p><img src=\"cid:image001.png@01DB2F1A\"/>";
        builder.Attachments.Add("Invoice.pdf", Encoding.ASCII.GetBytes("%PDF-1.4"), new ContentType("application", "pdf"));
        builder.Attachments.Add("Site_Photo.png", Png, new ContentType("image", "png"));
        return new MimeMessage { Subject = "Invoice", Body = builder.ToMessageBody() };
    }

    [Fact]
    public void Extract_Should_SkipSignatureLogo_When_HtmlReferencesItsContentId()
    {
        var attachments = EmailAttachments.Extract(EmailWithSignatureLogo());

        Assert.Equal(["Invoice.pdf", "Site_Photo.png"], attachments.Select(a => a.FileName));
    }

    [Fact]
    public void Extract_Should_ReturnAttachedEmailAsMessage_When_MessageIsAttached()
    {
        var inner = new MimeMessage { Subject = "Original notice", Body = new TextPart("plain") { Text = "Hearing moved." } };
        var builder = new BodyBuilder { TextBody = "Forwarding." };
        builder.Attachments.Add(new MessagePart { Message = inner });
        var outer = new MimeMessage { Subject = "Fwd", Body = builder.ToMessageBody() };

        var attachment = Assert.Single(EmailAttachments.Extract(outer));

        Assert.True(attachment.IsMessage);
        Assert.Equal("Original notice.eml", attachment.FileName);
    }

    [Fact]
    public void Extract_Should_ReturnNothing_When_EmailHasOnlyABody()
    {
        var message = new MimeMessage { Body = new TextPart("plain") { Text = "No attachments." } };

        Assert.Empty(EmailAttachments.Extract(message));
    }

    [Fact]
    public void ToWinAnsi_Should_MapSmartPunctuationToAscii_When_SubjectHasEnDashAndQuotes()
    {
        Assert.Equal("Holiday Notice - \"Tuesday\" ... it's", PdfText.ToWinAnsi("Holiday Notice – “Tuesday” … it’s"));
    }

    [Fact]
    public void Assemble_Should_LabelAndNestAttachments_When_EmailHasAttachments()
    {
        static byte[] OnePage(string title) => TextPdfRenderer.RenderText(title, "body");
        var members = new[]
        {
            new MemberPdf(1, "Letter.docx", MemberKind.Converted, OnePage("Letter")),
            new MemberPdf(2, "Invoice email.eml", MemberKind.Converted, OnePage("Email")),
            new MemberPdf(2, "Invoice.pdf", MemberKind.Pdf, OnePage("Invoice"), Label: "2.1", Level: 1),
            new MemberPdf(2, "Photo.png", MemberKind.Image, OnePage("Photo"), Label: "2.2", Level: 1),
            new MemberPdf(3, "Notes.txt", MemberKind.Text, OnePage("Notes"))
        };

        var bundle = PdfBundleAssembler.Assemble("Labels", members);

        Assert.Equal(["1", "2", "2.1", "2.2", "3"], bundle.Entries.Select(e => e.Label));
        Assert.Equal((4, 4), (bundle.Entries[2].StartPage, bundle.Entries[2].EndPage));
        using var pdf = PdfDocument.Open(bundle.Pdf);
        Assert.Contains("2.1", pdf.GetPage(1).Text);
        Assert.True(pdf.TryGetBookmarks(out var bookmarks));
        var email = bookmarks.Roots.Single(b => b.Title.StartsWith("2."));
        Assert.Equal(["2.1. Invoice.pdf", "2.2. Photo.png"], email.Children.Select(c => c.Title));
    }
}

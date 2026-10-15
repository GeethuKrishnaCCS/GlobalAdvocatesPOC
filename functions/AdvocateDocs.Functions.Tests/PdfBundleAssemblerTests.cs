using System.Text;
using AdvocateDocs.Functions.Bundles;
using UglyToad.PdfPig;

namespace AdvocateDocs.Functions.Tests;

public class PdfBundleAssemblerTests
{
    /// <summary>Text long enough to span the requested number of pages (TextPdfRenderer fits ~62 lines per A4 page).</summary>
    private static byte[] PdfWithPages(string title, int pages)
        => TextPdfRenderer.RenderText(title, string.Join('\n', Enumerable.Range(1, pages * 62 - 3).Select(i => $"line {i}")));

    [Fact]
    public void Assemble_Should_NumberMembersAfterIndex_When_SingleIndexPage()
    {
        var members = new[]
        {
            new MemberPdf(1, "First.docx", MemberKind.Converted, PdfWithPages("First", 2)),
            new MemberPdf(2, "Second.pdf", MemberKind.Pdf, PdfWithPages("Second", 3))
        };

        var bundle = PdfBundleAssembler.Assemble("Test bundle", members);

        Assert.Equal(1, bundle.IndexPageCount);
        Assert.Equal(6, bundle.PageCount);
        Assert.Equal((2, 3), (bundle.Entries[0].StartPage, bundle.Entries[0].EndPage));
        Assert.Equal((4, 6), (bundle.Entries[1].StartPage, bundle.Entries[1].EndPage));
        using var pdf = PdfDocument.Open(bundle.Pdf);
        Assert.Equal(6, pdf.NumberOfPages);
    }

    [Fact]
    public void Assemble_Should_StampEveryPageWithBundlePageNumber_When_Built()
    {
        var bundle = PdfBundleAssembler.Assemble("Stamp test", [new MemberPdf(1, "Only.txt", MemberKind.Text, PdfWithPages("Only", 2))]);

        using var pdf = PdfDocument.Open(bundle.Pdf);
        for (var p = 1; p <= pdf.NumberOfPages; p++)
        {
            Assert.Contains($"Bundle page {p}", pdf.GetPage(p).Text);
        }
    }

    [Fact]
    public void Assemble_Should_AddSecondIndexPage_When_MoreMembersThanFitOnOnePage()
    {
        var members = Enumerable.Range(1, PdfBundleAssembler.IndexRowsPerPage + 1)
            .Select(i => new MemberPdf(i, $"Doc {i}", MemberKind.Pdf, PdfWithPages($"Doc {i}", 1)))
            .ToList();

        var bundle = PdfBundleAssembler.Assemble("Large", members);

        Assert.Equal(2, bundle.IndexPageCount);
        Assert.Equal(3, bundle.Entries[0].StartPage);
        Assert.Equal(members.Count + 2, bundle.PageCount);
    }

    [Fact]
    public void Assemble_Should_ListMembersInOrder_When_InputUnordered()
    {
        var members = new[]
        {
            new MemberPdf(2, "B", MemberKind.Pdf, PdfWithPages("B", 1)),
            new MemberPdf(1, "A", MemberKind.Pdf, PdfWithPages("A", 1))
        };

        var bundle = PdfBundleAssembler.Assemble("Order", members);

        Assert.Equal(["A", "B"], bundle.Entries.Select(e => e.Title));
        using var pdf = PdfDocument.Open(bundle.Pdf);
        Assert.Contains("A", pdf.GetPage(2).Text);
    }

    [Fact]
    public void RenderEmail_Should_IncludeHeadersAndBody_When_PlainTextMessage()
    {
        var eml = Encoding.UTF8.GetBytes("From: a@x.test\r\nTo: b@x.test\r\nSubject: Hearing date\r\nDate: Mon, 1 Jan 2024 10:00:00 +0000\r\nContent-Type: text/plain\r\n\r\nThe hearing moved.");

        using var pdf = PdfDocument.Open(TextPdfRenderer.RenderEmail("mail.eml", eml));

        var text = pdf.GetPage(1).Text;
        Assert.Contains("Hearing date", text);
        Assert.Contains("The hearing moved.", text);
    }

    [Fact]
    public void ToWinAnsi_Should_ReplaceCharactersOutsideLatin1_When_Present()
        => Assert.Equal("Café ? ok", PdfText.ToWinAnsi("Café 中 ok"));
}

using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Core;
using UglyToad.PdfPig.Fonts.Standard14Fonts;
using UglyToad.PdfPig.Outline;
using UglyToad.PdfPig.Outline.Destinations;
using UglyToad.PdfPig.Writer;

namespace AdvocateDocs.Functions.Bundles;

/// <summary>
/// Builds the bundle PDF: index page(s) first, then every member's pages in order. Every page — index included —
/// is stamped "Bundle page N", and each member gets a bookmark at its first page.
/// </summary>
public static class PdfBundleAssembler
{
    public const int IndexRowsPerPage = 32;
    private const double Margin = 48;
    private const double StampSize = 9;
    private const double StampPadding = 3;
    private const int IndexTitleChars = 78;

    public static AssembledBundle Assemble(string bundleTitle, IReadOnlyList<MemberPdf> members)
    {
        var documents = members.OrderBy(m => m.Order).Select(m => (Member: m, Doc: PdfDocument.Open(m.Pdf))).ToList();
        try
        {
            var indexPages = Math.Max(1, (int)Math.Ceiling(documents.Count / (double)IndexRowsPerPage));
            var entries = new List<IndexEntry>(documents.Count);
            var next = indexPages + 1;
            foreach (var (member, doc) in documents)
            {
                entries.Add(new IndexEntry(member.Order, member.Title, member.Kind, next, next + doc.NumberOfPages - 1, member.Note));
                next += doc.NumberOfPages;
            }
            var pageCount = next - 1;

            using var builder = new PdfDocumentBuilder();
            var regular = builder.AddStandard14Font(Standard14Font.Helvetica);
            var bold = builder.AddStandard14Font(Standard14Font.HelveticaBold);

            WriteIndex(builder, bundleTitle, entries, indexPages, regular, bold);

            var bundlePage = indexPages;
            foreach (var (_, doc) in documents)
            {
                for (var p = 1; p <= doc.NumberOfPages; p++)
                {
                    bundlePage++;
                    Stamp(builder.AddPage(doc, p), bundlePage, regular);
                }
            }

            var bookmarks = new List<BookmarkNode>
            {
                new DocumentBookmarkNode("Index", 0, PageDestination(1), [])
            };
            bookmarks.AddRange(entries.Select(e =>
                new DocumentBookmarkNode($"{e.Order}. {PdfText.ToWinAnsi(e.Title)}", 0, PageDestination(e.StartPage), [])));
            builder.Bookmarks = new Bookmarks(bookmarks);

            return new AssembledBundle(builder.Build(), pageCount, indexPages, entries);
        }
        finally
        {
            documents.ForEach(d => d.Doc.Dispose());
        }
    }

    private static void WriteIndex(PdfDocumentBuilder builder, string bundleTitle, IReadOnlyList<IndexEntry> entries,
        int indexPages, PdfDocumentBuilder.AddedFont regular, PdfDocumentBuilder.AddedFont bold)
    {
        for (var pageIndex = 0; pageIndex < indexPages; pageIndex++)
        {
            var page = builder.AddPage(PageSize.A4);
            var top = page.PageSize.Height - Margin;
            var right = page.PageSize.Width - Margin;

            page.AddText(PdfText.Truncate(PdfText.ToWinAnsi($"Index - {bundleTitle}"), 70), 14, new PdfPoint(Margin, top), bold);
            var y = top - 28;
            page.AddText("No.", 9, new PdfPoint(Margin, y), bold);
            page.AddText("Document", 9, new PdfPoint(Margin + 34, y), bold);
            page.AddText("Bundle pages", 9, new PdfPoint(right - 64, y), bold);
            page.DrawLine(new PdfPoint(Margin, y - 4), new PdfPoint(right, y - 4), 0.5);
            y -= 18;

            foreach (var entry in entries.Skip(pageIndex * IndexRowsPerPage).Take(IndexRowsPerPage))
            {
                var pages = entry.StartPage == entry.EndPage ? $"{entry.StartPage}" : $"{entry.StartPage}-{entry.EndPage}";
                var title = PdfText.Truncate(PdfText.ToWinAnsi(entry.Title), IndexTitleChars)
                    + (entry.Kind == MemberKind.Placeholder ? " (not included)" : string.Empty);
                page.AddText($"{entry.Order}", 9, new PdfPoint(Margin, y), regular);
                page.AddText(title, 9, new PdfPoint(Margin + 34, y), regular);
                page.AddText(pages, 9, new PdfPoint(right - 64, y), regular);
                y -= 21;
            }
            Stamp(page, pageIndex + 1, regular);
        }
    }

    /// <summary>"Bundle page N" bottom-right on a white box, so it stays readable over page content.</summary>
    private static void Stamp(PdfPageBuilder page, int bundlePage, PdfDocumentBuilder.AddedFont font)
    {
        var text = $"Bundle page {bundlePage}";
        var size = page.PageSize;
        var letters = page.MeasureText(text, StampSize, new PdfPoint(0, 0), font);
        var width = letters.Count == 0 ? text.Length * StampSize * 0.5 : letters[^1].EndBaseLine.X - letters[0].StartBaseLine.X;
        var x = size.Right - 24 - width;
        var y = size.Bottom + 14;

        page.SetTextAndFillColor(255, 255, 255);
        page.DrawRectangle(new PdfPoint(x - StampPadding, y - StampPadding), width + 2 * StampPadding, StampSize + 2 * StampPadding, 0, fill: true);
        page.SetTextAndFillColor(0, 0, 0);
        page.AddText(text, StampSize, new PdfPoint(x, y), font);
        page.ResetColor();
    }

    private static ExplicitDestination PageDestination(int page)
        => new(page, ExplicitDestinationType.FitPage, ExplicitDestinationCoordinates.Empty);
}

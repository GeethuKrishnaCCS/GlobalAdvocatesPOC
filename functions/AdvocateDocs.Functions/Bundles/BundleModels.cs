namespace AdvocateDocs.Functions.Bundles;

/// <summary>How a member ended up in the bundle.</summary>
public static class MemberKind
{
    public const string Pdf = "pdf";
    public const string Converted = "converted";
    public const string Text = "text";
    public const string Placeholder = "placeholder";
}

/// <summary>One bundle member's PDF, in bundle order.</summary>
public sealed record MemberPdf(int Order, string Title, string Kind, byte[] Pdf, string? Note = null);

/// <summary>Index line: where a member sits in the bundle (bundle page numbers, 1-based, inclusive).</summary>
public sealed record IndexEntry(int Order, string Title, string Kind, int StartPage, int EndPage, string? Note);

/// <summary>Assembled bundle PDF and its index.</summary>
public sealed record AssembledBundle(byte[] Pdf, int PageCount, int IndexPageCount, IReadOnlyList<IndexEntry> Entries);

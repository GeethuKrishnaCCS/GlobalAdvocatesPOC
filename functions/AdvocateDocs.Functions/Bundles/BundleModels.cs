namespace AdvocateDocs.Functions.Bundles;

/// <summary>How a member ended up in the bundle.</summary>
public static class MemberKind
{
    public const string Pdf = "pdf";
    public const string Converted = "converted";
    public const string Text = "text";
    public const string Image = "image";
    /// <summary>Preview image of the first page only (no full conversion available, e.g. Visio).</summary>
    public const string Preview = "preview";
    public const string Placeholder = "placeholder";
}

/// <summary>
/// One index entry's PDF, in bundle order: a member (Level 0), or an email attachment following its email
/// (Level 1+, same Order, Label like "4.1").
/// </summary>
public sealed record MemberPdf(int Order, string Title, string Kind, byte[] Pdf, string? Note = null, string? Label = null, int Level = 0);

/// <summary>Index line: where an entry sits in the bundle (bundle page numbers, 1-based, inclusive).</summary>
public sealed record IndexEntry(int Order, string Title, string Kind, int StartPage, int EndPage, string? Note, string Label = "", int Level = 0);

/// <summary>Assembled bundle PDF and its index.</summary>
public sealed record AssembledBundle(byte[] Pdf, int PageCount, int IndexPageCount, IReadOnlyList<IndexEntry> Entries);

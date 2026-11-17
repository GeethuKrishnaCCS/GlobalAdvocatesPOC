using System.Text;

namespace AdvocateDocs.Functions.Bundles;

/// <summary>Text helpers for PDF's built-in (Standard 14) fonts, which only cover WinAnsi characters.</summary>
public static class PdfText
{
    private const char Replacement = '?';

    /// <summary>Typographic punctuation common in email subjects and file names, as plain ASCII (outside Latin-1).</summary>
    private static readonly Dictionary<char, string> Typographic = new()
    {
        ['–'] = "-", ['—'] = "-", ['‘'] = "'", ['’'] = "'", ['“'] = "\"", ['”'] = "\"",
        ['…'] = "...", ['•'] = "*"
    };

    /// <summary>Keeps printable Latin-1 characters; anything else becomes '?', so built-in fonts can draw it.</summary>
    public static string ToWinAnsi(string? value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return string.Empty;
        }
        var sb = new StringBuilder(value.Length);
        foreach (var ch in value)
        {
            if (Typographic.TryGetValue(ch, out var ascii))
            {
                sb.Append(ascii);
                continue;
            }
            sb.Append(ch is '\t' ? ' ' : ch >= ' ' && ch <= '~' || ch >= ' ' && ch <= 'ÿ' ? ch : Replacement);
        }
        return sb.ToString();
    }

    /// <summary>Truncates to a character budget with an ellipsis-like marker.</summary>
    public static string Truncate(string value, int maxChars)
        => value.Length <= maxChars ? value : value[..Math.Max(0, maxChars - 3)] + "...";

    /// <summary>Hard-wraps a line at a fixed width (monospace layout).</summary>
    public static IEnumerable<string> Wrap(string line, int width)
    {
        if (line.Length == 0)
        {
            yield return string.Empty;
            yield break;
        }
        for (var i = 0; i < line.Length; i += width)
        {
            yield return line.Substring(i, Math.Min(width, line.Length - i));
        }
    }
}

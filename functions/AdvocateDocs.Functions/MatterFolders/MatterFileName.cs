using System.Globalization;
using System.Text.RegularExpressions;

namespace AdvocateDocs.Functions.MatterFolders;

/// <summary>File name for a mail filed to a matter library: <c>yyyy-MM-dd subject.eml</c>.</summary>
public static partial class MatterFileName
{
    private const string NoSubject = "(no subject)";
    private const string Extension = ".eml";
    private const int MaxLength = 250;

    [GeneratedRegex("[\"*:<>?/\\\\|\\u0000-\\u001f]")]
    private static partial Regex IllegalChars();

    [GeneratedRegex("\\s+")]
    private static partial Regex Whitespace();

    /// <summary>Illegal characters become '-', whitespace collapses, the subject is trimmed of spaces and dots, and the name is capped at 250 characters.</summary>
    public static string For(DateTimeOffset receivedUtc, string? subject)
    {
        var clean = Whitespace().Replace(IllegalChars().Replace(subject ?? string.Empty, "-"), " ").Trim(' ', '.');
        if (clean.Length == 0)
        {
            clean = NoSubject;
        }
        var prefix = $"{receivedUtc.UtcDateTime.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)} ";
        var room = MaxLength - prefix.Length - Extension.Length;
        return prefix + (clean.Length > room ? clean[..room].TrimEnd() : clean) + Extension;
    }
}

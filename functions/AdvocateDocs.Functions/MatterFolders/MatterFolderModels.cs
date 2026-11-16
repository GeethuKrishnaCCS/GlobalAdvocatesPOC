namespace AdvocateDocs.Functions.MatterFolders;

/// <summary>One mail folder mapped to a library folder.</summary>
public sealed class MappedFolder
{
    public string MailFolderId { get; set; } = string.Empty;

    /// <summary>Library folder name; empty for the matter folder itself.</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>Drive path mail is filed under (empty = library root).</summary>
    public string DrivePath { get; set; } = string.Empty;

    /// <summary>Graph delta link of the last fully filed sweep; null = start from the beginning.</summary>
    public string? DeltaLink { get; set; }
}

/// <summary>Stored registration of one matter for one user (blob <c>{UserId}/{MatterRef}.json</c>).</summary>
public sealed class MatterFolderState
{
    public string UserId { get; set; } = string.Empty;
    public string UserPrincipalName { get; set; } = string.Empty;
    public string MatterRef { get; set; } = string.Empty;
    public string WebUrl { get; set; } = string.Empty;
    public string SiteId { get; set; } = string.Empty;
    public string DriveId { get; set; } = string.Empty;
    public List<MappedFolder> Folders { get; set; } = [];
    public int FiledCount { get; set; }
    public DateTimeOffset? LastSweepUtc { get; set; }
    public string? LastError { get; set; }
}

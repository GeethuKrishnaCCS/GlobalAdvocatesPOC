using AdvocateDocs.Functions.Graph;

namespace AdvocateDocs.Functions.MatterFolders;

/// <summary>Registers and unregisters a user's matter, acting as that user: mail folders in the mailbox plus the stored state the sweep works from.</summary>
public sealed class MatterFolderService(IUserGraph userGraph, MatterFolderStore store)
{
    public const string MattersFolderName = "Matters";

    /// <summary>Library folder the matter folder itself files to, when it exists.</summary>
    public const string MatterDrivePath = "Correspondence/Matter";

    /// <summary>
    /// Resolves the matter library, ensures <c>Matters/{matterRef}/{library folder}</c> exist in the user's mailbox and stores the
    /// state, keeping delta links of folders that still exist. Fails (not found) when the library does not exist.
    /// </summary>
    public async Task<Result<MatterFolderState>> RegisterAsync(string userId, string userPrincipalName, string matterRef,
        Uri webUrl, string documentsList, CancellationToken ct)
    {
        var (mail, graph) = userGraph.For(userId);
        var siteId = await graph.GetSiteIdAsync(webUrl, ct);
        var lists = await graph.GetListIdsByNameAsync(siteId, ct);
        if (!lists.TryGetValue(documentsList, out var listId))
        {
            return Result<MatterFolderState>.Fail($"List '{documentsList}' not found.");
        }
        var driveId = await graph.GetListDriveIdAsync(siteId, listId, ct);
        var libraryFolders = await graph.ListTopLevelFolderNamesAsync(driveId, ct);
        var matterPath = await graph.FolderExistsAsync(driveId, MatterDrivePath, ct) ? MatterDrivePath : string.Empty;
        var existing = (await store.GetAsync(userId, matterRef, ct))?.State;

        var targets = new List<(string Name, string DrivePath)> { (string.Empty, matterPath) };
        targets.AddRange(libraryFolders.Select(n => (n, n)));
        var resolved = await EnsureMailFoldersAsync(mail, userId, matterRef, targets, existing?.Folders ?? [], ct);

        var state = new MatterFolderState
        {
            UserId = userId, UserPrincipalName = userPrincipalName, MatterRef = matterRef, WebUrl = webUrl.ToString(),
            SiteId = siteId, DriveId = driveId, Folders = MergeFolders(existing?.Folders ?? [], resolved),
            FiledCount = existing?.FiledCount ?? 0, LastSweepUtc = existing?.LastSweepUtc, LastError = existing?.LastError
        };
        await store.SaveAsync(state, null, ct);
        return Result<MatterFolderState>.Ok(state);
    }

    /// <summary>
    /// Removes the stored state and deletes the matter's mail folders that hold no mail (leaves first, then the matter folder).
    /// Never deletes <c>Matters</c> or a folder with messages or subfolders.
    /// </summary>
    public async Task UnregisterAsync(string userId, string matterRef, CancellationToken ct)
    {
        var state = (await store.GetAsync(userId, matterRef, ct))?.State;
        if (state is null)
        {
            return;
        }
        var mail = userGraph.For(userId).Mail;
        foreach (var folder in state.Folders.OrderBy(f => f.Name.Length == 0))
        {
            var info = await mail.GetFolderAsync(userId, folder.MailFolderId, ct);
            if (info is { TotalItemCount: 0, ChildFolderCount: 0 })
            {
                await mail.DeleteFolderAsync(userId, folder.MailFolderId, ct);
            }
        }
        await store.DeleteAsync(userId, matterRef, ct);
    }

    /// <summary>The resolved folders, each keeping the delta link of the stored folder with the same name and mail folder id.</summary>
    public static List<MappedFolder> MergeFolders(IReadOnlyList<MappedFolder> existing, IReadOnlyList<MappedFolder> resolved)
        => resolved.Select(r => new MappedFolder
        {
            MailFolderId = r.MailFolderId, Name = r.Name, DrivePath = r.DrivePath,
            DeltaLink = existing.FirstOrDefault(e => e.Name == r.Name && e.MailFolderId == r.MailFolderId)?.DeltaLink
        }).ToList();

    private static async Task<List<MappedFolder>> EnsureMailFoldersAsync(GraphMailClient mail, string userId, string matterRef,
        IReadOnlyList<(string Name, string DrivePath)> targets, IReadOnlyList<MappedFolder> stored, CancellationToken ct)
    {
        var mattersId = await EnsureFolderAsync(mail, userId, null, MattersFolderName, null, ct);
        string? matterFolderId = null;
        var resolved = new List<MappedFolder>();
        foreach (var (name, drivePath) in targets)
        {
            var storedId = stored.FirstOrDefault(s => s.Name == name)?.MailFolderId;
            var id = name.Length == 0
                ? matterFolderId = await EnsureFolderAsync(mail, userId, mattersId, matterRef, storedId, ct)
                : await EnsureFolderAsync(mail, userId, matterFolderId, name, storedId, ct);
            resolved.Add(new MappedFolder { MailFolderId = id, Name = name, DrivePath = drivePath });
        }
        return resolved;
    }

    /// <summary>The stored folder when it still exists, else the folder with that name under the parent, else a new one.</summary>
    private static async Task<string> EnsureFolderAsync(GraphMailClient mail, string userId, string? parentId, string name, string? storedId, CancellationToken ct)
    {
        if (storedId is not null && await mail.GetFolderAsync(userId, storedId, ct) is not null)
        {
            return storedId;
        }
        return await mail.FindFolderAsync(userId, parentId, name, ct) ?? await mail.CreateFolderAsync(userId, parentId, name, ct);
    }
}

using System.Net;
using AdvocateDocs.Functions.Graph;
using Microsoft.Extensions.Logging;

namespace AdvocateDocs.Functions.MatterFolders;

/// <summary>
/// Files mail found in the registered matter mail folders to the matter library as .eml, as the user who pinned the matter, then tags it so it is not filed twice.
/// A folder's delta link only advances when every message of the round was filed.
/// </summary>
public sealed class MatterFolderSweeper(IUserGraph userGraph, MatterFolderStore store, ILogger<MatterFolderSweeper> logger)
{
    public const string FiledCategory = "Filed to matter";
    public const string FolderDeletedError = "Mail folder was deleted; reopen the add-in to recreate it.";
    private const int MaxErrorLength = 200;

    /// <summary>Sweeps every stored matter; one matter's failure never stops the others.</summary>
    public async Task RunAsync(CancellationToken ct)
    {
        await foreach (var stored in store.ListAsync(null, ct))
        {
            try
            {
                await SweepMatterAsync(stored.State, ct);
                // A lost race (the user re-registered or unregistered meanwhile) skips the write; filing is idempotent through the category.
                await store.SaveAsync(stored.State, stored.ETag, ct);
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                logger.LogError(ex, "Sweep failed for user {UserId} matter {MatterRef}", stored.State.UserId, stored.State.MatterRef);
            }
        }
    }

    /// <summary>Sweeps every mapped folder of one matter and updates its state (filed count, last sweep, last error, delta links).</summary>
    public async Task SweepMatterAsync(MatterFolderState state, CancellationToken ct)
    {
        var clients = userGraph.For(state.UserId);
        var filed = 0;
        string? error = null;
        foreach (var folder in state.Folders)
        {
            try
            {
                var (count, failure) = await SweepFolderAsync(clients, state, folder, ct);
                filed += count;
                error = failure ?? error;
            }
            catch (GraphRequestException ex) when (ex.Status == HttpStatusCode.NotFound)
            {
                folder.DeltaLink = null;
                error = FolderDeletedError;
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                error = ShortError(ex);
            }
        }
        state.FiledCount += filed;
        state.LastSweepUtc = DateTimeOffset.UtcNow;
        state.LastError = error;
        logger.LogInformation("Swept user {UserId} matter {MatterRef}: filed {Filed}, error {HasError}", state.UserId, state.MatterRef, filed, error is not null);
    }

    /// <summary>Filed count and, when any message failed, a short error (the delta link then stays where it was).</summary>
    private static async Task<(int Filed, string? Error)> SweepFolderAsync(UserGraphClients clients, MatterFolderState state, MappedFolder folder, CancellationToken ct)
    {
        var delta = await clients.Mail.GetDeltaAsync(state.UserId, folder.MailFolderId, folder.DeltaLink, ct);
        var filed = 0;
        string? failure = null;
        foreach (var message in delta.Messages.Where(m => !m.Categories.Contains(FiledCategory, StringComparer.OrdinalIgnoreCase)))
        {
            try
            {
                await FileAsync(clients, state, folder, message, ct);
                filed++;
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                failure ??= ShortError(ex);
            }
        }
        if (failure is null)
        {
            folder.DeltaLink = delta.DeltaLink ?? folder.DeltaLink;
        }
        return (filed, failure);
    }

    /// <summary>Uploads before tagging, so a failure in between files the message again rather than losing it.</summary>
    private static async Task FileAsync(UserGraphClients clients, MatterFolderState state, MappedFolder folder, MailMessageRef message, CancellationToken ct)
    {
        var mime = await clients.Mail.DownloadMimeAsync(state.UserId, message.Id, ct);
        var name = MatterFileName.For(message.ReceivedUtc ?? DateTimeOffset.UtcNow, message.Subject);
        var path = folder.DrivePath.Length == 0 ? name : $"{folder.DrivePath}/{name}";
        await clients.Sites.UploadAsync(state.DriveId, path, mime, ct, GraphSiteClient.ConflictRename);
        await clients.Mail.SetCategoriesAsync(state.UserId, message.Id, message.Categories.Append(FiledCategory), ct);
    }

    private static string ShortError(Exception ex) => ex.Message.Length <= MaxErrorLength ? ex.Message : ex.Message[..MaxErrorLength];
}

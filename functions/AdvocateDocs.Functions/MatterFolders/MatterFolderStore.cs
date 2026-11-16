using System.Runtime.CompilerServices;
using System.Text.Json;
using Azure;
using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;

namespace AdvocateDocs.Functions.MatterFolders;

/// <summary>A stored state with the blob ETag it was read at.</summary>
public sealed record StoredMatterState(MatterFolderState State, ETag ETag);

/// <summary>One JSON blob per user and matter in the Function's storage account (container <c>matter-folders</c>).</summary>
public sealed class MatterFolderStore(BlobServiceClient blobs)
{
    private const string ContainerName = "matter-folders";
    private static readonly JsonSerializerOptions Json = new(JsonSerializerDefaults.Web);

    public async Task<StoredMatterState?> GetAsync(string userId, string matterRef, CancellationToken ct)
        => await ReadAsync(await ContainerAsync(ct), BlobName(userId, matterRef), ct);

    /// <summary>All states, or only those of <paramref name="userId"/> when given.</summary>
    public async IAsyncEnumerable<StoredMatterState> ListAsync(string? userId, [EnumeratorCancellation] CancellationToken ct)
    {
        var container = await ContainerAsync(ct);
        await foreach (var blob in container.GetBlobsAsync(BlobTraits.None, BlobStates.None, prefix:userId is null ? null : $"{userId}/", cancellationToken: ct))
        {
            var stored = await ReadAsync(container, blob.Name, ct);
            if (stored is not null)
            {
                yield return stored;
            }
        }
    }

    /// <summary>Writes the state; with <paramref name="ifMatch"/> only when the blob is unchanged. False when that condition failed.</summary>
    public async Task<bool> SaveAsync(MatterFolderState state, ETag? ifMatch, CancellationToken ct)
    {
        var blob = (await ContainerAsync(ct)).GetBlobClient(BlobName(state.UserId, state.MatterRef));
        var content = BinaryData.FromObjectAsJson(state, Json);
        if (ifMatch is null)
        {
            await blob.UploadAsync(content, overwrite: true, ct);
            return true;
        }
        try
        {
            await blob.UploadAsync(content, new BlobUploadOptions { Conditions = new BlobRequestConditions { IfMatch = ifMatch } }, ct);
            return true;
        }
        catch (RequestFailedException ex) when (ex.Status is 404 or 412)
        {
            return false;
        }
    }

    public async Task DeleteAsync(string userId, string matterRef, CancellationToken ct)
        => await (await ContainerAsync(ct)).DeleteBlobIfExistsAsync(BlobName(userId, matterRef), cancellationToken: ct);

    private static string BlobName(string userId, string matterRef) => $"{userId}/{matterRef}.json";

    private async Task<BlobContainerClient> ContainerAsync(CancellationToken ct)
    {
        var container = blobs.GetBlobContainerClient(ContainerName);
        await container.CreateIfNotExistsAsync(cancellationToken: ct);
        return container;
    }

    private static async Task<StoredMatterState?> ReadAsync(BlobContainerClient container, string name, CancellationToken ct)
    {
        try
        {
            var result = (await container.GetBlobClient(name).DownloadContentAsync(ct)).Value;
            return new StoredMatterState(result.Content.ToObjectFromJson<MatterFolderState>(Json)!, result.Details.ETag);
        }
        catch (RequestFailedException ex) when (ex.Status == 404)
        {
            return null;
        }
    }
}

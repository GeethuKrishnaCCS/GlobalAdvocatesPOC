using Azure.Storage.Blobs;

namespace AdvocateDocs.Functions.Bundles;

/// <summary>
/// Scratch space for a build (converted members, assembled PDF) in the Function's storage account, so large
/// payloads never pass through orchestration state. Each build uses its own prefix and is removed at the end.
/// </summary>
public sealed class BundleWorkStore(BlobServiceClient blobs)
{
    private const string ContainerName = "bundle-work";

    public async Task PutAsync(string name, byte[] content, CancellationToken ct)
    {
        var container = blobs.GetBlobContainerClient(ContainerName);
        await container.CreateIfNotExistsAsync(cancellationToken: ct);
        await container.GetBlobClient(name).UploadAsync(new BinaryData(content), overwrite: true, ct);
    }

    public async Task<byte[]> GetAsync(string name, CancellationToken ct)
        => (await blobs.GetBlobContainerClient(ContainerName).GetBlobClient(name).DownloadContentAsync(ct)).Value.Content.ToArray();

    public async Task DeletePrefixAsync(string prefix, CancellationToken ct)
    {
        var container = blobs.GetBlobContainerClient(ContainerName);
        await foreach (var blob in container.GetBlobsAsync(Azure.Storage.Blobs.Models.BlobTraits.None, Azure.Storage.Blobs.Models.BlobStates.None, prefix, ct))
        {
            await container.DeleteBlobIfExistsAsync(blob.Name, cancellationToken: ct);
        }
    }
}

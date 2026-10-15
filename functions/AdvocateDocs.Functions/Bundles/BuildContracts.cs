namespace AdvocateDocs.Functions.Bundles;

/// <summary>URL names (not titles) of the matter site's lists/libraries involved in a build.</summary>
public sealed record MatterLists(string Documents, string Bundles, string Registry, string Items, string Manifests);

/// <summary>What the SPFx command posts to start a build.</summary>
public sealed record BuildRequest(string WebUrl, string BundleId, MatterLists Lists);

/// <summary>Orchestration input: the request plus who asked.</summary>
public sealed record BuildInput(BuildRequest Request, string RequestedBy);

/// <summary>A member to include, in bundle order.</summary>
public sealed record PlannedMember(int Order, int SourceItemId, string SourceUniqueId, string Title);

/// <summary>Resolved ids and members for one build.</summary>
public sealed record BundlePlan(
    string SiteId,
    string DocumentsListId,
    string BundlesListId,
    string RegistryListId,
    string ManifestsListId,
    int RegistryItemId,
    string BundleId,
    string Title,
    IReadOnlyList<PlannedMember> Members);

/// <summary>Convert activity input.</summary>
public sealed record ConvertInput(string WorkPrefix, string SiteId, string DocumentsListId, PlannedMember Member);

/// <summary>A converted member parked in blob storage.</summary>
public sealed record ConvertedMember(int Order, string Title, string Kind, string BlobName, string SourceUniqueId, int SourceItemId, string? Note);

/// <summary>Assemble activity input.</summary>
public sealed record AssembleInput(string WorkPrefix, string Title, IReadOnlyList<ConvertedMember> Members);

/// <summary>Assembled bundle parked in blob storage.</summary>
public sealed record AssembleOutput(string BlobName, int PageCount, int IndexPageCount, IReadOnlyList<IndexEntry> Entries);

/// <summary>Publish activity input.</summary>
public sealed record PublishInput(string WorkPrefix, BuildInput Build, BundlePlan Plan, IReadOnlyList<ConvertedMember> Members, AssembleOutput Assembled);

/// <summary>Build outcome returned to the caller.</summary>
public sealed record BuildResult(string PdfUrl, int Version, int PageCount, int Documents, int Placeholders);

/// <summary>Progress shown while the orchestration runs.</summary>
public sealed record BuildProgress(string Stage, int Done, int Total);

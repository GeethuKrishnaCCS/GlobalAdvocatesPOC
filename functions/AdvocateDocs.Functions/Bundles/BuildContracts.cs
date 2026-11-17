namespace AdvocateDocs.Functions.Bundles;

/// <summary>URL names (not titles) of the matter site's lists/libraries involved in a build.</summary>
/// <param name="Conversion">Hidden library for temporary copies Graph converts to PDF (email attachments); optional.</param>
public sealed record MatterLists(string Documents, string Bundles, string Registry, string Items, string Manifests, string? Conversion = null);

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
    IReadOnlyList<PlannedMember> Members,
    string? ConversionListId = null);

/// <summary>Convert activity input.</summary>
public sealed record ConvertInput(string WorkPrefix, string SiteId, string DocumentsListId, PlannedMember Member, string? ConversionListId = null);

/// <summary>An email attachment converted to PDF and parked in blob storage; follows its email in the bundle.</summary>
/// <param name="Label">Index number, e.g. "4.1" (attachment 1 of member 4) or "4.2.1" (inside an attached email).</param>
public sealed record ConvertedPart(string Label, int Level, string Title, string Kind, string BlobName, string? Note);

/// <summary>A converted member parked in blob storage, plus its email attachments (if any) in order.</summary>
public sealed record ConvertedMember(int Order, string Title, string Kind, string BlobName, string SourceUniqueId, int SourceItemId, string? Note,
    IReadOnlyList<ConvertedPart>? Attachments = null);

/// <summary>Assemble activity input.</summary>
public sealed record AssembleInput(string WorkPrefix, string Title, IReadOnlyList<ConvertedMember> Members);

/// <summary>Assembled bundle parked in blob storage.</summary>
public sealed record AssembleOutput(string BlobName, int PageCount, int IndexPageCount, IReadOnlyList<IndexEntry> Entries);

/// <summary>Publish activity input.</summary>
public sealed record PublishInput(string WorkPrefix, BuildInput Build, BundlePlan Plan, IReadOnlyList<ConvertedMember> Members, AssembleOutput Assembled);

/// <summary>Build outcome returned to the caller.</summary>
/// <param name="Placeholders">Index entries (members or attachments) that could not be included.</param>
/// <param name="Attachments">Email attachments added as their own index entries.</param>
public sealed record BuildResult(string PdfUrl, int Version, int PageCount, int Documents, int Placeholders, int Attachments = 0);

/// <summary>Progress shown while the orchestration runs.</summary>
public sealed record BuildProgress(string Stage, int Done, int Total, string? Error = null);

namespace AdvocateDocs.Functions.Auth;

/// <summary>Entra ID settings for validating callers' access tokens (app settings section <c>Entra</c>).</summary>
public sealed class EntraOptions
{
    public const string SectionName = "Entra";

    /// <summary>Directory (tenant) id the tokens must be issued by.</summary>
    public string TenantId { get; init; } = string.Empty;

    /// <summary>API app registration client id (v2 tokens carry it as the audience).</summary>
    public string ClientId { get; init; } = string.Empty;

    /// <summary>Application ID URI (v1 tokens carry it as the audience), e.g. <c>api://{clientId}</c>.</summary>
    public string Audience { get; init; } = string.Empty;

    /// <summary>Delegated scope callers must hold.</summary>
    public string RequiredScope { get; init; } = "user_impersonation";

    /// <summary>
    /// Client id of the user-assigned managed identity the API app registration trusts as its credential
    /// (federated identity credential); needed only to act on behalf of users (matter mail folders).
    /// </summary>
    public string FederatedIdentityClientId { get; init; } = string.Empty;
}

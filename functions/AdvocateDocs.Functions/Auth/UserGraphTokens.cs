using System.Net.Http.Headers;
using Azure;
using Azure.Core;
using Azure.Identity;
using Azure.Storage.Blobs;
using Microsoft.Extensions.Options;
using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Extensibility;

namespace AdvocateDocs.Functions.Auth;

/// <summary>The user's stored sign-in is missing or no longer accepted; they must call the API again from the add-in.</summary>
public sealed class SignInRequiredException(Exception inner) : Exception(UserGraphTokens.SignInRequiredMessage, inner);

/// <summary>
/// Graph tokens on behalf of a user that outlive the request (on-behalf-of, long-running): the timer files mail as the
/// user, so SharePoint and Exchange enforce the user's own access. One token cache blob per user (container
/// <c>matter-folders-tokens</c>); the app authenticates with a managed identity as federated credential, no secret.
/// </summary>
public sealed class UserGraphTokens
{
    public const string SignInRequiredMessage = "Sign-in expired. Open the add-in and choose Retry to resume filing.";
    private const string ContainerName = "matter-folders-tokens";
    private const string TokenExchangeScope = "api://AzureADTokenExchange/.default";
    private static readonly string[] GraphScopes = ["https://graph.microsoft.com/.default"];
    private static readonly TimeSpan ExpiryMargin = TimeSpan.FromMinutes(5);

    private readonly ILongRunningWebApi _app;
    private readonly BlobServiceClient _blobs;
    private BlobContainerClient? _container;

    public UserGraphTokens(IOptions<EntraOptions> entra, BlobServiceClient blobs)
    {
        var options = entra.Value;
        if (string.IsNullOrWhiteSpace(options.FederatedIdentityClientId))
        {
            throw new InvalidOperationException("Entra:FederatedIdentityClientId is required to act on behalf of users.");
        }
        _blobs = blobs;
        var identity = new ManagedIdentityCredential(ManagedIdentityId.FromUserAssignedClientId(options.FederatedIdentityClientId));
        var app = ConfidentialClientApplicationBuilder.Create(options.ClientId)
            .WithAuthority($"https://login.microsoftonline.com/{options.TenantId}")
            .WithClientAssertion(async (AssertionRequestOptions request) =>
                (await identity.GetTokenAsync(new TokenRequestContext([TokenExchangeScope]), request.CancellationToken)).Token)
            .Build();
        // shortcut: token cache blobs rely on storage encryption at rest and a private container, add envelope encryption (Key Vault key) before a customer rollout.
        app.UserTokenCache.SetBeforeAccessAsync(LoadCacheAsync);
        app.UserTokenCache.SetAfterAccessAsync(SaveCacheAsync);
        _app = (ILongRunningWebApi)app;
    }

    /// <summary>The bearer token of an Authorization header value.</summary>
    public static string? BearerOf(string? authorizationHeader)
        => AuthenticationHeaderValue.TryParse(authorizationHeader, out var header) && header.Scheme.Equals("Bearer", StringComparison.OrdinalIgnoreCase)
            ? header.Parameter
            : null;

    /// <summary>Exchanges the caller's access token for Graph tokens and stores them for later use without the user.</summary>
    public async Task BeginAsync(string userId, string userAccessToken, CancellationToken ct)
    {
        var key = SessionKey(userId);
        try
        {
            await _app.InitiateLongRunningProcessInWebApi(GraphScopes, userAccessToken, ref key).ExecuteAsync(ct);
        }
        catch (MsalUiRequiredException ex)
        {
            throw new SignInRequiredException(ex);
        }
    }

    /// <summary>Removes the user's stored tokens.</summary>
    public Task EndAsync(string userId, CancellationToken ct) => _app.StopLongRunningProcessInWebApiAsync(SessionKey(userId), ct);

    /// <summary>A credential for Graph as this user; throws <see cref="SignInRequiredException"/> when the stored sign-in is gone.</summary>
    public TokenCredential For(string userId) => new UserCredential(_app, SessionKey(userId));

    /// <summary>The session key names a blob: only a directory object id (GUID) is accepted.</summary>
    private static string SessionKey(string userId)
        => Guid.TryParse(userId, out var id) ? id.ToString("D") : throw new ArgumentException("userId must be a GUID.", nameof(userId));

    private async Task LoadCacheAsync(TokenCacheNotificationArgs args)
    {
        byte[]? cache = null;
        try
        {
            cache = (await (await BlobAsync(args)).DownloadContentAsync(args.CancellationToken)).Value.Content.ToArray();
        }
        catch (RequestFailedException ex) when (ex.Status == 404)
        {
            // No stored sign-in for this user.
        }
        args.TokenCache.DeserializeMsalV3(cache, shouldClearExistingCache: true);
    }

    private async Task SaveCacheAsync(TokenCacheNotificationArgs args)
    {
        if (!args.HasStateChanged)
        {
            return;
        }
        var blob = await BlobAsync(args);
        if (args.HasTokens)
        {
            await blob.UploadAsync(BinaryData.FromBytes(args.TokenCache.SerializeMsalV3()), overwrite: true, args.CancellationToken);
        }
        else
        {
            await blob.DeleteIfExistsAsync(cancellationToken: args.CancellationToken);
        }
    }

    private async Task<BlobClient> BlobAsync(TokenCacheNotificationArgs args)
    {
        if (_container is null)
        {
            var container = _blobs.GetBlobContainerClient(ContainerName);
            await container.CreateIfNotExistsAsync(cancellationToken: args.CancellationToken);
            _container = container;
        }
        return _container.GetBlobClient($"{SessionKey(args.SuggestedCacheKey)}.bin");
    }

    /// <summary>Keeps the access token until shortly before it expires, so one sweep reads the cache blob once.</summary>
    private sealed class UserCredential(ILongRunningWebApi app, string sessionKey) : TokenCredential
    {
        private AccessToken? _token;

        public override AccessToken GetToken(TokenRequestContext requestContext, CancellationToken cancellationToken)
            => throw new NotSupportedException("Use GetTokenAsync.");

        public override async ValueTask<AccessToken> GetTokenAsync(TokenRequestContext requestContext, CancellationToken cancellationToken)
        {
            if (_token is { } cached && cached.ExpiresOn - DateTimeOffset.UtcNow > ExpiryMargin)
            {
                return cached;
            }
            try
            {
                var result = await app.AcquireTokenInLongRunningProcess(requestContext.Scopes, sessionKey).ExecuteAsync(cancellationToken);
                _token = new AccessToken(result.AccessToken, result.ExpiresOn);
                return _token.Value;
            }
            catch (MsalException ex) when (ex is MsalUiRequiredException || ex.ErrorCode == MsalError.OboCacheKeyNotInCacheError)
            {
                throw new SignInRequiredException(ex);
            }
        }
    }
}

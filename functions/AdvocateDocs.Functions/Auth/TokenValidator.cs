using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.IdentityModel.Protocols;
using Microsoft.IdentityModel.Protocols.OpenIdConnect;
using Microsoft.IdentityModel.Tokens;

namespace AdvocateDocs.Functions.Auth;

/// <summary>Validates Entra ID bearer tokens issued to the AdvocateDocs API (signature, issuer, audience, lifetime, scope).</summary>
public interface ITokenValidator
{
    /// <summary>Returns the caller's principal, or a failure reason when the token is missing or invalid.</summary>
    Task<Result<ClaimsPrincipal>> ValidateAsync(string? authorizationHeader, CancellationToken cancellationToken);
}

/// <inheritdoc />
public sealed class TokenValidator : ITokenValidator
{
    private const string BearerPrefix = "Bearer ";
    private const string ScopeClaim = "http://schemas.microsoft.com/identity/claims/scope";
    private const string ShortScopeClaim = "scp";

    private readonly EntraOptions _options;
    private readonly ILogger<TokenValidator> _logger;
    private readonly ConfigurationManager<OpenIdConnectConfiguration> _metadata;
    private readonly JwtSecurityTokenHandler _handler = new() { MapInboundClaims = false };

    public TokenValidator(IOptions<EntraOptions> options, ILogger<TokenValidator> logger)
    {
        _options = options.Value;
        _logger = logger;
        // v2.0 metadata: signing keys are shared by v1 and v2 tokens of the tenant.
        _metadata = new ConfigurationManager<OpenIdConnectConfiguration>(
            $"https://login.microsoftonline.com/{_options.TenantId}/v2.0/.well-known/openid-configuration",
            new OpenIdConnectConfigurationRetriever());
    }

    public async Task<Result<ClaimsPrincipal>> ValidateAsync(string? authorizationHeader, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(authorizationHeader) || !authorizationHeader.StartsWith(BearerPrefix, StringComparison.OrdinalIgnoreCase))
        {
            return Result<ClaimsPrincipal>.Fail("Missing bearer token.");
        }

        var token = authorizationHeader[BearerPrefix.Length..].Trim();
        var config = await _metadata.GetConfigurationAsync(cancellationToken);
        var parameters = new TokenValidationParameters
        {
            ValidIssuers =
            [
                $"https://login.microsoftonline.com/{_options.TenantId}/v2.0",
                $"https://sts.windows.net/{_options.TenantId}/"
            ],
            ValidAudiences = [_options.ClientId, _options.Audience],
            IssuerSigningKeys = config.SigningKeys,
            ValidateLifetime = true,
            ClockSkew = TimeSpan.FromMinutes(2)
        };

        try
        {
            var principal = _handler.ValidateToken(token, parameters, out _);
            var scopes = (principal.FindFirst(ShortScopeClaim) ?? principal.FindFirst(ScopeClaim))?.Value.Split(' ') ?? [];
            return scopes.Contains(_options.RequiredScope, StringComparer.OrdinalIgnoreCase)
                ? Result<ClaimsPrincipal>.Ok(principal)
                : Result<ClaimsPrincipal>.Fail($"Token lacks scope '{_options.RequiredScope}'.");
        }
        // Malformed (undecodable) tokens surface as ArgumentException, not SecurityTokenException.
        catch (Exception ex) when (ex is SecurityTokenException or ArgumentException)
        {
            _logger.LogWarning("Rejected token: {Reason}", ex.Message);
            return Result<ClaimsPrincipal>.Fail("Invalid token.");
        }
    }
}

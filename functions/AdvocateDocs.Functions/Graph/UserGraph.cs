using AdvocateDocs.Functions.Auth;

namespace AdvocateDocs.Functions.Graph;

/// <summary>Graph clients that act as one user.</summary>
public sealed record UserGraphClients(GraphMailClient Mail, GraphSiteClient Sites);

/// <summary>Graph on behalf of a user (delegated), as opposed to the Function's own identity.</summary>
public interface IUserGraph
{
    /// <summary>Clients for the user's mailbox and the sites the user can reach.</summary>
    UserGraphClients For(string userId);
}

/// <inheritdoc />
public sealed class UserGraph(IHttpClientFactory http, UserGraphTokens tokens) : IUserGraph
{
    public const string HttpClientName = nameof(UserGraph);

    public UserGraphClients For(string userId)
    {
        var credential = tokens.For(userId);
        var client = http.CreateClient(HttpClientName);
        return new UserGraphClients(new GraphMailClient(client, credential), new GraphSiteClient(client, credential));
    }
}

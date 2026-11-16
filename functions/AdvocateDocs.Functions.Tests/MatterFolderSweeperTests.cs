using System.Net;
using System.Text;
using AdvocateDocs.Functions.Auth;
using AdvocateDocs.Functions.Graph;
using AdvocateDocs.Functions.MatterFolders;
using Azure.Core;
using Microsoft.Extensions.Logging.Abstractions;

namespace AdvocateDocs.Functions.Tests;

public class MatterFolderSweeperTests
{
    private const string OldLink = "https://graph.microsoft.com/v1.0/users/u1/mailFolders/f1/messages/delta?$deltatoken=old";
    private const string NewLink = "https://graph.microsoft.com/v1.0/users/u1/mailFolders/f1/messages/delta?$deltatoken=new";

    private sealed class FakeCredential : TokenCredential
    {
        public override AccessToken GetToken(TokenRequestContext requestContext, CancellationToken cancellationToken) => new("t", DateTimeOffset.UtcNow.AddHours(1));
        public override ValueTask<AccessToken> GetTokenAsync(TokenRequestContext requestContext, CancellationToken cancellationToken) => new(GetToken(requestContext, cancellationToken));
    }

    private sealed class FakeUserGraph(UserGraphClients clients) : IUserGraph
    {
        public UserGraphClients For(string userId) => clients;
    }

    private sealed class FakeHandler(Func<HttpRequestMessage, string, HttpResponseMessage> respond) : HttpMessageHandler
    {
        public List<string> Calls { get; } = [];
        public List<string> Bodies { get; } = [];

        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            var body = request.Content is null ? string.Empty : await request.Content.ReadAsStringAsync(cancellationToken);
            Calls.Add($"{request.Method} {request.RequestUri}");
            Bodies.Add(body);
            return respond(request, body);
        }
    }

    private static HttpResponseMessage Json(string json, HttpStatusCode status = HttpStatusCode.OK)
        => new(status) { Content = new StringContent(json, Encoding.UTF8, "application/json") };

    private static string Delta(string categories) =>
        $$"""{"value":[{"id":"m1","subject":"Hi: there","receivedDateTime":"2026-11-13T10:00:00Z","categories":{{categories}}}],"@odata.deltaLink":"{{NewLink}}"}""";

    private static HttpResponseMessage Route(HttpRequestMessage r, string delta, bool failUpload = false)
    {
        var url = r.RequestUri!.ToString();
        if (url.Contains("/messages/delta") || url.Contains("deltatoken")) return Json(delta);
        if (url.EndsWith("/messages/m1/$value")) return new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent([1, 2, 3]) };
        if (url.Contains("createUploadSession")) return failUpload ? Json("""{"error":{"message":"boom"}}""", HttpStatusCode.InternalServerError) : Json("""{"uploadUrl":"https://upload.test/u"}""");
        if (url.StartsWith("https://upload.test/")) return Json("""{"id":"i1"}""", HttpStatusCode.Created);
        if (url.Contains("/drives/d1/items/i1")) return Json("""{"id":"i1","name":"n.eml","webUrl":"https://x/n.eml","listItem":{"id":"5"}}""");
        if (r.Method == HttpMethod.Patch) return Json("{}");
        return Json("{}", HttpStatusCode.NotFound);
    }

    private static readonly Uri ApiBase = new("https://func.test/api/");

    private static MatterFolderService Service(IUserGraph userGraph)
        => new(userGraph, new MatterFolderStore(new Azure.Storage.Blobs.BlobServiceClient("UseDevelopmentStorage=true")), ApiBase);

    private static MatterFolderSweeper Sweeper(IUserGraph userGraph)
        => new(userGraph, new MatterFolderStore(new Azure.Storage.Blobs.BlobServiceClient("UseDevelopmentStorage=true")), Service(userGraph), NullLogger<MatterFolderSweeper>.Instance);

    private static (MatterFolderSweeper Sweeper, FakeHandler Handler, MatterFolderState State) Arrange(Func<HttpRequestMessage, string, HttpResponseMessage> respond)
    {
        var handler = new FakeHandler(respond);
        var http = new HttpClient(handler);
        var userGraph = new FakeUserGraph(new UserGraphClients(new GraphMailClient(http, new FakeCredential()), new GraphSiteClient(http, new FakeCredential())));
        var sweeper = Sweeper(userGraph);
        var state = new MatterFolderState
        {
            UserId = "u1", MatterRef = "M-1", DriveId = "d1",
            Folders = [new MappedFolder { MailFolderId = "f1", Name = "", DrivePath = "Correspondence/Matter", DeltaLink = OldLink }]
        };
        return (sweeper, handler, state);
    }

    [Fact]
    public async Task SweepMatterAsync_Should_UploadWithRenameThenTag_When_MessageUnfiled()
    {
        var (sweeper, handler, state) = Arrange((r, _) => Route(r, Delta("[\"Red category\"]")));

        await sweeper.SweepMatterAsync(state, CancellationToken.None);

        var upload = handler.Calls.FindIndex(c => c.Contains("createUploadSession"));
        var tag = handler.Calls.FindIndex(c => c.StartsWith("PATCH "));
        Assert.InRange(upload, 0, tag - 1);
        Assert.Contains("\"rename\"", handler.Bodies[upload]);
        Assert.Contains("Correspondence/Matter/2026-11-13 Hi- there.eml", handler.Calls[upload]);
        Assert.Contains("Filed to matter", handler.Bodies[tag]);
        Assert.Contains("Red category", handler.Bodies[tag]);
        Assert.Equal(NewLink, state.Folders[0].DeltaLink);
        Assert.Equal(1, state.FiledCount);
        Assert.Null(state.LastError);
    }

    [Fact]
    public async Task SweepMatterAsync_Should_SkipMessage_When_AlreadyTagged()
    {
        var (sweeper, handler, state) = Arrange((r, _) => Route(r, Delta("[\"Filed to matter\"]")));

        await sweeper.SweepMatterAsync(state, CancellationToken.None);

        Assert.DoesNotContain(handler.Calls, c => c.Contains("createUploadSession") || c.StartsWith("PATCH "));
        Assert.Equal(NewLink, state.Folders[0].DeltaLink);
        Assert.Equal(0, state.FiledCount);
    }

    [Fact]
    public async Task SweepMatterAsync_Should_KeepDeltaLinkAndSetError_When_UploadFails()
    {
        var (sweeper, handler, state) = Arrange((r, _) => Route(r, Delta("[]"), failUpload: true));

        await sweeper.SweepMatterAsync(state, CancellationToken.None);

        Assert.Equal(OldLink, state.Folders[0].DeltaLink);
        Assert.NotNull(state.LastError);
        Assert.Equal(0, state.FiledCount);
        Assert.DoesNotContain(handler.Calls, c => c.StartsWith("PATCH "));
    }

    [Fact]
    public async Task SweepMatterAsync_Should_RecordDeletedFolder_When_DeltaReturns404()
    {
        var (sweeper, _, state) = Arrange((_, _) => Json("""{"error":{"message":"gone"}}""", HttpStatusCode.NotFound));

        await sweeper.SweepMatterAsync(state, CancellationToken.None);

        Assert.Equal(MatterFolderSweeper.FolderDeletedError, state.LastError);
        Assert.Null(state.Folders[0].DeltaLink);
    }

    private sealed class ExpiredSignIn : TokenCredential
    {
        public override AccessToken GetToken(TokenRequestContext requestContext, CancellationToken cancellationToken) => throw new SignInRequiredException(new InvalidOperationException("expired"));
        public override ValueTask<AccessToken> GetTokenAsync(TokenRequestContext requestContext, CancellationToken cancellationToken) => throw new SignInRequiredException(new InvalidOperationException("expired"));
    }

    [Fact]
    public async Task SweepMatterAsync_Should_AskForSignInAndKeepDeltaLink_When_StoredSignInExpired()
    {
        var http = new HttpClient(new FakeHandler((_, _) => Json("{}")));
        var sweeper = Sweeper(new FakeUserGraph(new UserGraphClients(new GraphMailClient(http, new ExpiredSignIn()), new GraphSiteClient(http, new ExpiredSignIn()))));
        var state = new MatterFolderState { UserId = "u1", MatterRef = "M-1", DriveId = "d1", Folders = [new MappedFolder { MailFolderId = "f1", Name = "", DeltaLink = OldLink }] };

        await sweeper.SweepMatterAsync(state, CancellationToken.None);

        Assert.Equal(UserGraphTokens.SignInRequiredMessage, state.LastError);
        Assert.Equal(OldLink, state.Folders[0].DeltaLink);
    }

    [Fact]
    public async Task EnsureSubscriptionsAsync_Should_RenewLiveAndRecreateGone_When_FoldersHaveSubscriptions()
    {
        var handler = new FakeHandler((r, _) => r.Method == HttpMethod.Patch
            ? Json("{}", r.RequestUri!.ToString().EndsWith("/gone") ? HttpStatusCode.NotFound : HttpStatusCode.OK)
            : Json("""{"id":"new"}""", HttpStatusCode.Created));
        var http = new HttpClient(handler);
        var service = Service(new FakeUserGraph(new UserGraphClients(new GraphMailClient(http, new FakeCredential()), new GraphSiteClient(http, new FakeCredential()))));
        var state = new MatterFolderState
        {
            UserId = "u1", MatterRef = "M 1",
            Folders = [new MappedFolder { MailFolderId = "f1", SubscriptionId = "live" }, new MappedFolder { MailFolderId = "f2", SubscriptionId = "gone" }, new MappedFolder { MailFolderId = "f3" }]
        };

        await service.EnsureSubscriptionsAsync(state, CancellationToken.None);

        Assert.Equal(["live", "new", "new"], state.Folders.Select(f => f.SubscriptionId));
        Assert.All(state.Folders, f => Assert.NotNull(f.SubscriptionExpiresUtc));
        var created = handler.Bodies[handler.Calls.FindIndex(c => c.StartsWith("POST "))];
        Assert.Contains("https://func.test/api/matter-folders/notify?u=u1&m=M%201", created.Replace("\\u0026", "&"));
        Assert.Contains(state.NotificationSecret!, created);
        Assert.Contains("mailFolders('f2')/messages", created.Replace("\\u0027", "'"));
    }

    [Fact]
    public void MergeFolders_Should_KeepDeltaLinkOnlyOfSurvivingFolders_When_Reregistered()
    {
        var existing = new List<MappedFolder>
        {
            new() { MailFolderId = "a", Name = "", DeltaLink = "link-a" },
            new() { MailFolderId = "b", Name = "Letters", DeltaLink = "link-b" },
            new() { MailFolderId = "c", Name = "Gone", DeltaLink = "link-c" }
        };
        var resolved = new List<MappedFolder>
        {
            new() { MailFolderId = "a", Name = "", DrivePath = "Correspondence/Matter" },
            new() { MailFolderId = "b2", Name = "Letters", DrivePath = "Letters" },
            new() { MailFolderId = "d", Name = "New", DrivePath = "New" }
        };

        var merged = MatterFolderService.MergeFolders(existing, resolved);

        Assert.Equal("link-a", merged[0].DeltaLink);
        Assert.Equal("Correspondence/Matter", merged[0].DrivePath);
        Assert.Null(merged[1].DeltaLink);
        Assert.Null(merged[2].DeltaLink);
        Assert.Equal(3, merged.Count);
    }
}

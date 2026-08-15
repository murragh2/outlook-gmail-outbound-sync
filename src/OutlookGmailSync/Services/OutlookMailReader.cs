using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions.Authentication;
using DeltaGetResponse = Microsoft.Graph.Me.MailFolders.Item.Messages.Delta.DeltaGetResponse;

namespace OutlookGmailSync.Services;

public interface IOutlookMailReader
{
    Task<(IReadOnlyList<Message> Messages, string? NewDeltaLink)> GetNewSentItemsAsync(string? deltaLink, CancellationToken ct);
}

public class OutlookMailReader : IOutlookMailReader
{
    private readonly IGraphTokenService _tokenService;
    private readonly ILogger<OutlookMailReader>? _logger;

    public OutlookMailReader(IGraphTokenService tokenService, ILogger<OutlookMailReader>? logger = null)
    {
        _tokenService = tokenService;
        _logger = logger;
    }

    public async Task<(IReadOnlyList<Message> Messages, string? NewDeltaLink)> GetNewSentItemsAsync(string? deltaLink, CancellationToken ct)
    {
        var graphClient = CreateGraphClient();
        var messages = new List<Message>();
        string? nextDeltaLink = null;

        try
        {
            if (string.IsNullOrEmpty(deltaLink))
            {
                _logger?.LogInformation("[GRAPH READ] Requesting fresh delta token for SentItems folder...");
                var response = await graphClient.Me.MailFolders["SentItems"].Messages.Delta.GetAsDeltaGetResponseAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Select = new[] { "id", "internetMessageId", "subject", "sentDateTime", "toRecipients", "ccRecipients", "bccRecipients" };
                }, ct);

                await PageThroughResults(graphClient, response, messages, (link) => nextDeltaLink = link, ct);
            }
            else
            {
                _logger?.LogInformation("[GRAPH READ] Querying Graph SentItems delta endpoint with existing token...");
                var response = await graphClient.Me.MailFolders["SentItems"].Messages.Delta.WithUrl(deltaLink).GetAsDeltaGetResponseAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Select = new[] { "id", "internetMessageId", "subject", "sentDateTime", "toRecipients", "ccRecipients", "bccRecipients" };
                }, ct);

                await PageThroughResults(graphClient, response, messages, (link) => nextDeltaLink = link, ct);
            }
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError ex) when (ex.ResponseStatusCode == 410)
        {
            _logger?.LogWarning("[GRAPH READ] Delta token expired (HTTP 410 Gone). Re-establishing new delta link...");
            var response = await graphClient.Me.MailFolders["SentItems"].Messages.Delta.GetAsDeltaGetResponseAsync(requestConfiguration =>
            {
                requestConfiguration.QueryParameters.Select = new[] { "id", "internetMessageId", "subject", "sentDateTime", "toRecipients", "ccRecipients", "bccRecipients" };
            }, ct);
            
            await PageThroughResults(graphClient, response, messages, (link) => nextDeltaLink = link, ct);
        }

        _logger?.LogInformation("[GRAPH READ COMPLETE] Retrieved {Count} message(s). Has new DeltaLink: {HasLink}", messages.Count, !string.IsNullOrEmpty(nextDeltaLink));
        return (messages, nextDeltaLink);
    }

    private async Task PageThroughResults(GraphServiceClient client, DeltaGetResponse? response, List<Message> messages, Action<string?> setDeltaLink, CancellationToken ct)
    {
        if (response == null) return;
        var pageIterator = Microsoft.Graph.PageIterator<Message, DeltaGetResponse>
            .CreatePageIterator(client, response, (m) =>
            {
                messages.Add(m);
                return true;
            },
            (req) =>
            {
                return req;
            });

        await pageIterator.IterateAsync(ct);

        if (pageIterator.Deltalink != null)
        {
            setDeltaLink(pageIterator.Deltalink);
        }
    }

    private GraphServiceClient CreateGraphClient()
    {
        var authProvider = new BaseBearerTokenAuthenticationProvider(new TokenProvider(_tokenService));
        return new GraphServiceClient(authProvider);
    }

    private class TokenProvider : IAccessTokenProvider
    {
        private readonly IGraphTokenService _tokenService;

        public TokenProvider(IGraphTokenService tokenService)
        {
            _tokenService = tokenService;
        }

        public AllowedHostsValidator AllowedHostsValidator { get; } = new AllowedHostsValidator();

        public async Task<string> GetAuthorizationTokenAsync(Uri uri, Dictionary<string, object>? additionalAuthenticationContext = null, CancellationToken cancellationToken = default)
        {
            return await _tokenService.GetAccessTokenAsync(cancellationToken);
        }
    }
}

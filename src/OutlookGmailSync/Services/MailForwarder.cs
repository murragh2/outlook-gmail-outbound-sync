using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Graph.Me.Messages.Item.Forward;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions.Authentication;
using OutlookGmailSync.Configuration;

namespace OutlookGmailSync.Services;

public interface IMailForwarder
{
    Task<bool> ForwardAsync(Message message, CancellationToken ct);
}

public class MailForwarder : IMailForwarder
{
    private readonly IGraphTokenService _tokenService;
    private readonly SyncOptions _syncOptions;
    private readonly ILogger<MailForwarder> _logger;

    public MailForwarder(IGraphTokenService tokenService, IOptions<SyncOptions> syncOptions, ILogger<MailForwarder> logger)
    {
        _tokenService = tokenService;
        _syncOptions = syncOptions.Value;
        _logger = logger;
    }

    public async Task<bool> ForwardAsync(Message message, CancellationToken ct)
    {
        if (IsLoop(message))
        {
            _logger.LogInformation("Skipping message {Id} because it already contains the Gmail address (loop guard).", message.Id);
            return false;
        }

        var graphClient = CreateGraphClient();
        var toRecipients = new List<Recipient>
        {
            new Recipient
            {
                EmailAddress = new EmailAddress
                {
                    Address = _syncOptions.GmailAddress
                }
            }
        };

        var forwardBody = new ForwardPostRequestBody
        {
            ToRecipients = toRecipients,
            Comment = ""
        };

        try
        {
            await graphClient.Me.Messages[message.Id].Forward.PostAsync(forwardBody, cancellationToken: ct);
            return true;
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError ex) when (ex.ResponseStatusCode == 429)
        {
            _logger.LogWarning("Throttled while forwarding message {Id}.", message.Id);
            throw;
        }
    }

    private bool IsLoop(Message message)
    {
        var targetEmail = _syncOptions.GmailAddress;

        bool CheckRecipients(IEnumerable<Recipient>? recipients)
        {
            if (recipients == null) return false;
            return recipients.Any(r => string.Equals(r.EmailAddress?.Address, targetEmail, StringComparison.OrdinalIgnoreCase));
        }

        return CheckRecipients(message.ToRecipients) || CheckRecipients(message.CcRecipients) || CheckRecipients(message.BccRecipients);
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

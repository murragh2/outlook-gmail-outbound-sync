using FluentAssertions;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.Graph.Models;
using Moq;
using OutlookGmailSync.Configuration;
using OutlookGmailSync.Services;
using Xunit;

namespace OutlookGmailSync.Tests;

public class MailForwarderTests
{
    private readonly Mock<IGraphTokenService> _tokenServiceMock;
    private readonly IOptions<SyncOptions> _syncOptions;
    private readonly Mock<ILogger<MailForwarder>> _loggerMock;
    private readonly MailForwarder _forwarder;

    public MailForwarderTests()
    {
        _tokenServiceMock = new Mock<IGraphTokenService>();
        _syncOptions = Options.Create(new SyncOptions { GmailAddress = "target@gmail.com" });
        _loggerMock = new Mock<ILogger<MailForwarder>>();
        
        _forwarder = new MailForwarder(_tokenServiceMock.Object, _syncOptions, _loggerMock.Object);
    }

    [Fact]
    public async Task ForwardAsync_WithGmailInToRecipients_ReturnsFalse()
    {
        var message = new Message
        {
            Id = "msg-1",
            ToRecipients = new List<Recipient> { new Recipient { EmailAddress = new EmailAddress { Address = "target@gmail.com" } } }
        };

        var result = await _forwarder.ForwardAsync(message, CancellationToken.None);

        result.Should().BeFalse();
    }

    [Fact]
    public async Task ForwardAsync_WithGmailInCcRecipients_ReturnsFalse()
    {
        var message = new Message
        {
            Id = "msg-2",
            CcRecipients = new List<Recipient> { new Recipient { EmailAddress = new EmailAddress { Address = "TARGET@gmail.com" } } }
        };

        var result = await _forwarder.ForwardAsync(message, CancellationToken.None);

        result.Should().BeFalse();
    }
}

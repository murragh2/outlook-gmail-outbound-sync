using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Microsoft.Graph.Models;
using Moq;
using OutlookGmailSync.Functions;
using OutlookGmailSync.Models;
using OutlookGmailSync.Services;
using Xunit;

namespace OutlookGmailSync.Tests;

public class SyncTimerFunctionTests
{
    private readonly Mock<ISyncStateStore> _syncStateStoreMock;
    private readonly Mock<IOutlookMailReader> _mailReaderMock;
    private readonly Mock<IMailForwarder> _mailForwarderMock;
    private readonly Mock<IDeduplicationService> _dedupServiceMock;
    private readonly Mock<ILogger<SyncTimerFunction>> _loggerMock;
    private readonly SyncTimerFunction _function;

    public SyncTimerFunctionTests()
    {
        _syncStateStoreMock = new Mock<ISyncStateStore>();
        _mailReaderMock = new Mock<IOutlookMailReader>();
        _mailForwarderMock = new Mock<IMailForwarder>();
        _dedupServiceMock = new Mock<IDeduplicationService>();
        _loggerMock = new Mock<ILogger<SyncTimerFunction>>();

        _function = new SyncTimerFunction(
            _syncStateStoreMock.Object,
            _mailReaderMock.Object,
            _mailForwarderMock.Object,
            _dedupServiceMock.Object,
            _loggerMock.Object);
    }

    [Fact]
    public async Task Run_InitialSync_SkipsForwardingAndSavesDelta()
    {
        var state = new SyncState { IsInitialSyncComplete = false };
        _syncStateStoreMock.Setup(s => s.GetStateAsync(It.IsAny<CancellationToken>())).ReturnsAsync(state);
        
        var messages = new List<Message> { new Message() };
        _mailReaderMock.Setup(r => r.GetNewSentItemsAsync(null, It.IsAny<CancellationToken>()))
            .ReturnsAsync((messages, "new-delta-link"));

        await _function.Run(null!, CancellationToken.None);

        _mailForwarderMock.Verify(f => f.ForwardAsync(It.IsAny<Message>(), It.IsAny<CancellationToken>()), Times.Never);
        _syncStateStoreMock.Verify(s => s.SaveStateAsync(It.Is<SyncState>(x => x.IsInitialSyncComplete && x.DeltaLink == "new-delta-link"), It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task Run_NormalSync_ForwardsAndDeduplicates()
    {
        var state = new SyncState { IsInitialSyncComplete = true, DeltaLink = "old-delta" };
        _syncStateStoreMock.Setup(s => s.GetStateAsync(It.IsAny<CancellationToken>())).ReturnsAsync(state);
        
        var msg1 = new Message { InternetMessageId = "1", SentDateTime = DateTimeOffset.UtcNow };
        var msg2 = new Message { InternetMessageId = "2", SentDateTime = DateTimeOffset.UtcNow };
        var messages = new List<Message> { msg1, msg2 };
        
        _mailReaderMock.Setup(r => r.GetNewSentItemsAsync("old-delta", It.IsAny<CancellationToken>()))
            .ReturnsAsync((messages, "new-delta-link"));

        _dedupServiceMock.Setup(d => d.HasBeenForwardedAsync("1", It.IsAny<DateTimeOffset>(), It.IsAny<CancellationToken>())).ReturnsAsync(true);
        _dedupServiceMock.Setup(d => d.HasBeenForwardedAsync("2", It.IsAny<DateTimeOffset>(), It.IsAny<CancellationToken>())).ReturnsAsync(false);
        _mailForwarderMock.Setup(f => f.ForwardAsync(msg2, It.IsAny<CancellationToken>())).ReturnsAsync(true);

        await _function.Run(null!, CancellationToken.None);

        _mailForwarderMock.Verify(f => f.ForwardAsync(msg1, It.IsAny<CancellationToken>()), Times.Never);
        _mailForwarderMock.Verify(f => f.ForwardAsync(msg2, It.IsAny<CancellationToken>()), Times.Once);
        _dedupServiceMock.Verify(d => d.RecordForwardedAsync("2", It.IsAny<string>(), It.IsAny<DateTimeOffset>(), It.IsAny<CancellationToken>()), Times.Once);
    }
}

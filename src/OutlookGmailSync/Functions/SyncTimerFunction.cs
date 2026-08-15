using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using OutlookGmailSync.Models;
using OutlookGmailSync.Services;

namespace OutlookGmailSync.Functions;

public class SyncTimerFunction
{
    private readonly ISyncStateStore _syncStateStore;
    private readonly IOutlookMailReader _mailReader;
    private readonly IMailForwarder _mailForwarder;
    private readonly IDeduplicationService _dedupService;
    private readonly ILogger<SyncTimerFunction> _logger;

    public SyncTimerFunction(
        ISyncStateStore syncStateStore,
        IOutlookMailReader mailReader,
        IMailForwarder mailForwarder,
        IDeduplicationService dedupService,
        ILogger<SyncTimerFunction> logger)
    {
        _syncStateStore = syncStateStore;
        _mailReader = mailReader;
        _mailForwarder = mailForwarder;
        _dedupService = dedupService;
        _logger = logger;
    }

    [Function("SyncTimerFunction")]
    public async Task Run([TimerTrigger("%Sync:CronSchedule%")] TimerInfo timerInfo, CancellationToken ct)
    {
        _logger.LogInformation("SyncTimerFunction triggered at: {Time}", DateTimeOffset.UtcNow);

        try
        {
            var syncState = await _syncStateStore.GetStateAsync(ct) ?? new SyncState();
            
            var (messages, newDeltaLink) = await _mailReader.GetNewSentItemsAsync(syncState.DeltaLink, ct);

            if (!syncState.IsInitialSyncComplete)
            {
                _logger.LogInformation("Initial sync. Found {Count} messages. Storing delta link and skipping forwarding.", messages.Count);
                syncState.IsInitialSyncComplete = true;
                syncState.DeltaLink = newDeltaLink;
                syncState.LastSyncTime = DateTimeOffset.UtcNow;
                await _syncStateStore.SaveStateAsync(syncState, ct);
                return;
            }

            int forwarded = 0;
            int skippedDupes = 0;
            int skippedLoop = 0;

            foreach (var message in messages)
            {
                var internetMessageId = message.InternetMessageId ?? message.Id;
                if (string.IsNullOrEmpty(internetMessageId)) continue;
                var subject = message.Subject ?? "No Subject";
                var sentDateTime = message.SentDateTime ?? DateTimeOffset.UtcNow;

                bool hasBeenForwarded = await _dedupService.HasBeenForwardedAsync(internetMessageId, sentDateTime, ct);
                
                if (hasBeenForwarded)
                {
                    skippedDupes++;
                    continue;
                }

                bool wasForwarded = await _mailForwarder.ForwardAsync(message, ct);

                if (wasForwarded)
                {
                    await _dedupService.RecordForwardedAsync(internetMessageId, subject, sentDateTime, ct);
                    forwarded++;
                }
                else
                {
                    skippedLoop++;
                }
            }

            syncState.DeltaLink = newDeltaLink;
            syncState.LastSyncTime = DateTimeOffset.UtcNow;
            await _syncStateStore.SaveStateAsync(syncState, ct);

            _logger.LogInformation("Sync completed. Found: {Count}, Forwarded: {Forwarded}, Skipped (Dupes): {SkippedDupes}, Skipped (Loop): {SkippedLoop}", 
                messages.Count, forwarded, skippedDupes, skippedLoop);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "An error occurred during sync execution.");
            throw;
        }
    }
}

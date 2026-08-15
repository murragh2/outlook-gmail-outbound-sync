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
        var startTime = DateTimeOffset.UtcNow;
        _logger.LogInformation("================================================================================");
        _logger.LogInformation("[SYNC START] Outlook -> Gmail Sync Cycle Started at {Time} UTC", startTime);
        _logger.LogInformation("================================================================================");

        try
        {
            var syncState = await _syncStateStore.GetStateAsync(ct) ?? new SyncState();
            bool isInitial = !syncState.IsInitialSyncComplete || string.IsNullOrEmpty(syncState.DeltaLink);

            if (isInitial)
            {
                _logger.LogInformation("[SYNC MODE] Initial sync setup. Establishing initial delta watermark...");
            }
            else
            {
                _logger.LogInformation("[SYNC MODE] Incremental sync. Using existing delta link (Last sync: {LastSync})", syncState.LastSyncTime);
            }

            var (messages, newDeltaLink) = await _mailReader.GetNewSentItemsAsync(syncState.DeltaLink, ct);

            _logger.LogInformation("[GRAPH DELTA RESULT] Fetched {Count} new/modified message(s) from SentItems folder.", messages.Count);

            if (isInitial)
            {
                _logger.LogInformation("[INITIAL SYNC COMPLETE] Stored initial delta link. Skipping historical messages ({Count} items).", messages.Count);
                syncState.IsInitialSyncComplete = true;
                syncState.DeltaLink = newDeltaLink;
                syncState.LastSyncTime = DateTimeOffset.UtcNow;
                await _syncStateStore.SaveStateAsync(syncState, ct);
                return;
            }

            int forwarded = 0;
            int skippedDupes = 0;
            int skippedLoop = 0;

            for (int i = 0; i < messages.Count; i++)
            {
                var message = messages[i];
                var internetMessageId = message.InternetMessageId ?? message.Id;
                if (string.IsNullOrEmpty(internetMessageId))
                {
                    _logger.LogWarning("[MSG #{Index}] Skipping message with null or empty ID.", i + 1);
                    continue;
                }

                var subject = message.Subject ?? "(No Subject)";
                var sentDateTime = message.SentDateTime ?? DateTimeOffset.UtcNow;

                _logger.LogInformation("--- Processing [{Index}/{Total}] Subject: '{Subject}' | Sent: {SentTime} | MsgId: {MsgId} ---",
                    i + 1, messages.Count, subject, sentDateTime, internetMessageId);

                bool hasBeenForwarded = await _dedupService.HasBeenForwardedAsync(internetMessageId, sentDateTime, ct);
                if (hasBeenForwarded)
                {
                    _logger.LogInformation("[SKIP: DEDUPLICATED] MsgId '{MsgId}' was already forwarded. Skipping.", internetMessageId);
                    skippedDupes++;
                    continue;
                }

                bool wasForwarded = await _mailForwarder.ForwardAsync(message, ct);

                if (wasForwarded)
                {
                    await _dedupService.RecordForwardedAsync(internetMessageId, subject, sentDateTime, ct);
                    _logger.LogInformation("[SUCCESS: FORWARDED] MsgId '{MsgId}' ('{Subject}') forwarded and recorded in dedup table.", internetMessageId, subject);
                    forwarded++;
                }
                else
                {
                    _logger.LogInformation("[SKIP: LOOP GUARD] MsgId '{MsgId}' contains target Gmail recipient. Skipping.", internetMessageId);
                    skippedLoop++;
                }
            }

            syncState.DeltaLink = newDeltaLink;
            syncState.LastSyncTime = DateTimeOffset.UtcNow;
            await _syncStateStore.SaveStateAsync(syncState, ct);

            var duration = DateTimeOffset.UtcNow - startTime;
            _logger.LogInformation("================================================================================");
            _logger.LogInformation("[SYNC SUMMARY] Duration: {Duration}s | Total Delta: {Total} | Forwarded: {Forwarded} | Skipped (Dupes): {Dupes} | Skipped (Loop Guard): {Loop}",
                Math.Round(duration.TotalSeconds, 2), messages.Count, forwarded, skippedDupes, skippedLoop);
            _logger.LogInformation("================================================================================");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "[SYNC FAILED] An error occurred during sync cycle execution.");
            throw;
        }
    }
}

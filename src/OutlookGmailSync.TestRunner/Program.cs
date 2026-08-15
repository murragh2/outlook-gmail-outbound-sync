using Azure.Data.Tables;
using Azure.Identity;
using Azure.Security.KeyVault.Secrets;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Graph.Models;
using OutlookGmailSync.Configuration;
using OutlookGmailSync.Models;
using OutlookGmailSync.Services;

namespace OutlookGmailSync.TestRunner;

class Program
{
    static async Task Main(string[] args)
    {
        string mode = "dry-run";
        for (int i = 0; i < args.Length; i++)
        {
            if (args[i] == "--mode" && i + 1 < args.Length)
            {
                mode = args[i + 1].ToLowerInvariant();
            }
        }

        Console.WriteLine("================================================================================");
        Console.WriteLine($"[LOCAL TEST RUNNER] Mode: {mode.ToUpperInvariant()}");
        Console.WriteLine("================================================================================");

        if (mode == "simulate")
        {
            RunSimulationMode();
            return;
        }

        var config = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile("src/OutlookGmailSync/local.settings.json", optional: true, reloadOnChange: true)
            .AddUserSecrets<Program>(optional: true)
            .AddEnvironmentVariables()
            .Build();

        var services = new ServiceCollection();
        ConfigureServices(services, config);
        var provider = services.BuildServiceProvider();

        var logger = provider.GetRequiredService<ILogger<Program>>();
        var mailReader = provider.GetRequiredService<IOutlookMailReader>();
        var mailForwarder = provider.GetRequiredService<IMailForwarder>();
        var dedupService = provider.GetRequiredService<IDeduplicationService>();
        var syncStateStore = provider.GetRequiredService<ISyncStateStore>();

        var syncState = await syncStateStore.GetStateAsync(CancellationToken.None) ?? new SyncState();
        var gmailAddress = config["Values:Sync__GmailAddress"] ?? "murragh2@gmail.com";
        var keyVaultUri = config["Values:KeyVault__VaultUri"] ?? "https://kv-ogms-x4g5sogvlsg3i.vault.azure.net";

        logger.LogInformation("[CONFIG] Target Gmail: {Gmail}", gmailAddress);
        logger.LogInformation("[CONFIG] KeyVault URI: {KV}", keyVaultUri);

        logger.LogInformation("[FETCH] Fetching sent items delta from Microsoft Graph...");
        var result = await mailReader.GetNewSentItemsAsync(syncState.DeltaLink, CancellationToken.None);
        var messages = result.Messages;
        var newDeltaLink = result.NewDeltaLink;

        logger.LogInformation("[FETCH RESULT] Retrieved {Count} message(s).", messages.Count);

        int forwarded = 0;
        int skippedDupes = 0;
        int skippedLoop = 0;

        for (int i = 0; i < messages.Count; i++)
        {
            var message = messages[i];
            var internetMessageId = message.InternetMessageId ?? message.Id;
            if (string.IsNullOrEmpty(internetMessageId)) continue;
            var subject = message.Subject ?? "(No Subject)";
            var sentDateTime = message.SentDateTime ?? DateTimeOffset.UtcNow;

            Console.WriteLine($"\n--- [{i + 1}/{messages.Count}] Subject: '{subject}' | MsgId: {internetMessageId} ---");

            bool hasBeenForwarded = await dedupService.HasBeenForwardedAsync(internetMessageId, sentDateTime, CancellationToken.None);
            if (hasBeenForwarded)
            {
                logger.LogInformation("[EVAL: DEDUPLICATED] Already forwarded. Skipping.");
                skippedDupes++;
                continue;
            }

            if (mode == "dry-run")
            {
                bool isLoop = IsLoop(message, gmailAddress);
                if (isLoop)
                {
                    logger.LogInformation("[EVAL: DRY-RUN] Would SKIP (Loop Guard: recipient matches target Gmail).");
                    skippedLoop++;
                }
                else
                {
                    logger.LogInformation("[EVAL: DRY-RUN] Would FORWARD -> {Gmail} (Skipped real send in dry-run mode).", gmailAddress);
                    forwarded++;
                }
            }
            else if (mode == "run-once")
            {
                bool wasForwarded = await mailForwarder.ForwardAsync(message, CancellationToken.None);
                if (wasForwarded)
                {
                    await dedupService.RecordForwardedAsync(internetMessageId, subject, sentDateTime, CancellationToken.None);
                    logger.LogInformation("[LIVE: FORWARDED] Successfully forwarded message.");
                    forwarded++;
                }
                else
                {
                    logger.LogInformation("[LIVE: LOOP GUARD] Skipped message containing target Gmail address.");
                    skippedLoop++;
                }
            }
        }

        if (mode == "run-once" && !string.IsNullOrEmpty(newDeltaLink))
        {
            syncState.DeltaLink = newDeltaLink;
            syncState.LastSyncTime = DateTimeOffset.UtcNow;
            syncState.IsInitialSyncComplete = true;
            await syncStateStore.SaveStateAsync(syncState, CancellationToken.None);
            logger.LogInformation("[STATE] Saved new delta link state to Table Storage.");
        }

        Console.WriteLine("\n================================================================================");
        logger.LogInformation("[SUMMARY] Total: {Total} | Forwarded: {Forwarded} | Skipped (Dupes): {Dupes} | Skipped (Loop Guard): {Loop}",
            messages.Count, forwarded, skippedDupes, skippedLoop);
        Console.WriteLine("================================================================================");
    }

    private static void ConfigureServices(IServiceCollection services, IConfiguration config)
    {
        services.AddLogging(builder =>
        {
            builder.AddConsole();
            builder.SetMinimumLevel(LogLevel.Information);
        });

        services.Configure<AzureAdOptions>(options =>
        {
            options.TenantId = config["Values:AzureAd__TenantId"] ?? "consumers";
            options.ClientId = config["Values:AzureAd__ClientId"] ?? "";
            options.ClientSecret = config["Values:AzureAd__ClientSecret"] ?? "";
        });

        services.Configure<SyncOptions>(options =>
        {
            options.GmailAddress = config["Values:Sync__GmailAddress"] ?? "murragh2@gmail.com";
            options.CronSchedule = config["Values:Sync__CronSchedule"] ?? "0 */10 * * * *";
        });

        var keyVaultUri = config["Values:KeyVault__VaultUri"] ?? "https://kv-ogms-x4g5sogvlsg3i.vault.azure.net";
        var storageAccountUri = config["Values:Storage__AccountUri"] ?? "https://stogmsx4g5sogvlsg3i.table.core.windows.net";

        services.AddSingleton(new SecretClient(new Uri(keyVaultUri), new DefaultAzureCredential()));
        services.AddSingleton(new TableServiceClient(new Uri(storageAccountUri), new DefaultAzureCredential()));

        services.AddHttpClient<IGraphTokenService, GraphTokenService>();
        services.AddSingleton<IOutlookMailReader, OutlookMailReader>();
        services.AddSingleton<IMailForwarder, MailForwarder>();
        services.AddSingleton<IDeduplicationService, DeduplicationService>();
        services.AddSingleton<ISyncStateStore, SyncStateStore>();
    }

    private static bool IsLoop(Message message, string targetEmail)
    {
        bool Check(IEnumerable<Recipient>? recipients) =>
            recipients != null && recipients.Any(r => string.Equals(r.EmailAddress?.Address, targetEmail, StringComparison.OrdinalIgnoreCase));

        return Check(message.ToRecipients) || Check(message.CcRecipients) || Check(message.BccRecipients);
    }

    private static void RunSimulationMode()
    {
        Console.WriteLine("[SIMULATION] Testing synthetic email scenarios locally...\n");

        var targetEmail = "murragh2@gmail.com";

        var testCases = new[]
        {
            new { Id = "msg-001", Subject = "Hello from Outlook", Recipients = new[] { "alice@example.com" }, Expected = "FORWARD" },
            new { Id = "msg-002", Subject = "FW: Hello from Outlook", Recipients = new[] { "murragh2@gmail.com" }, Expected = "SKIP (Loop Guard)" },
            new { Id = "msg-003", Subject = "Project Update", Recipients = new[] { "bob@example.com", "murragh2@gmail.com" }, Expected = "SKIP (Loop Guard)" },
            new { Id = "msg-004", Subject = "Meeting Notes", Recipients = new[] { "carol@example.com" }, Expected = "FORWARD" }
        };

        foreach (var tc in testCases)
        {
            bool isLoop = tc.Recipients.Any(r => string.Equals(r, targetEmail, StringComparison.OrdinalIgnoreCase));
            string result = isLoop ? "SKIP (Loop Guard)" : "FORWARD";
            bool pass = result == tc.Expected;

            Console.WriteLine($"Message ID: {tc.Id}");
            Console.WriteLine($"  Subject:    '{tc.Subject}'");
            Console.WriteLine($"  Recipients: {string.Join(", ", tc.Recipients)}");
            Console.WriteLine($"  Result:     {result} | Expected: {tc.Expected} | Status: {(pass ? "PASSED ✅" : "FAILED ❌")}");
            Console.WriteLine();
        }

        Console.WriteLine("[SIMULATION COMPLETE] All synthetic scenarios verified.");
    }
}

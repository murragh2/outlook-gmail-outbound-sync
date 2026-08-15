using Azure.Data.Tables;
using Azure.Identity;
using Azure.Security.KeyVault.Secrets;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using NUnit.Framework;
using OutlookGmailSync.Configuration;
using OutlookGmailSync.Functions;
using OutlookGmailSync.Services;

namespace OutlookGmailSync.Tests;

[TestFixture]
[Category("LocalIntegration")]
public class LocalSyncIntegrationTests
{
    private IServiceProvider? _serviceProvider;
    private IConfiguration? _configuration;

    [SetUp]
    public void SetUp()
    {
        _configuration = new ConfigurationBuilder()
            .SetBasePath(AppContext.BaseDirectory)
            .AddJsonFile("local.settings.json", optional: true)
            .AddUserSecrets<LocalSyncIntegrationTests>(optional: true)
            .AddEnvironmentVariables()
            .Build();

        var localToken = _configuration["AzureAd:GraphAccessToken"];

        if (string.IsNullOrEmpty(localToken))
        {
            Assert.Ignore(
                "\n================================================================================\n" +
                "[INVALID SETUP ❌] Local Graph Access Token is missing from .NET User Secrets.\n" +
                "================================================================================\n" +
                "Please acquire a fresh 1-hour access token by running:\n" +
                "  dotnet run --project src/OutlookGmailSync.AuthCli -- --client-id 06d76858-fa3e-48c2-8b98-3f3d167efa6b --save-local\n" +
                "================================================================================");
            return;
        }

        var services = new ServiceCollection();

        services.AddLogging(builder =>
        {
            builder.AddConsole();
            builder.SetMinimumLevel(LogLevel.Information);
        });

        services.Configure<AzureAdOptions>(options =>
        {
            options.TenantId = _configuration["Values:AzureAd__TenantId"] ?? "consumers";
            options.ClientId = _configuration["Values:AzureAd__ClientId"] ?? "06d76858-fa3e-48c2-8b98-3f3d167efa6b";
            options.GraphAccessToken = localToken;
        });

        services.Configure<SyncOptions>(options =>
        {
            options.GmailAddress = _configuration["Values:Sync__GmailAddress"] ?? "murragh2@gmail.com";
            options.CronSchedule = _configuration["Values:Sync__CronSchedule"] ?? "0 */10 * * * *";
        });

        var keyVaultUri = _configuration["Values:KeyVault__VaultUri"] ?? "https://kv-ogms-x4g5sogvlsg3i.vault.azure.net";
        var storageAccountUri = _configuration["Values:Storage__AccountUri"] ?? "https://stogmsx4g5sogvlsg3i.table.core.windows.net";

        services.AddSingleton(new SecretClient(new Uri(keyVaultUri), new DefaultAzureCredential()));
        services.AddSingleton(new TableServiceClient(new Uri(storageAccountUri), new DefaultAzureCredential()));

        services.AddHttpClient<IGraphTokenService, GraphTokenService>();
        services.AddSingleton<IOutlookMailReader, OutlookMailReader>();
        services.AddSingleton<IMailForwarder, MailForwarder>();
        services.AddSingleton<IDeduplicationService, DeduplicationService>();
        services.AddSingleton<ISyncStateStore, SyncStateStore>();
        services.AddSingleton<SyncTimerFunction>();

        _serviceProvider = services.BuildServiceProvider();
    }

    [Test]
    public async Task ExecuteLocalSync_WithUserSecretsToken_ExecutesCleanly()
    {
        Assert.That(_serviceProvider, Is.Not.Null, "Service provider should be initialized when User Secrets token is present.");
        var function = _serviceProvider!.GetRequiredService<SyncTimerFunction>();

        var timerInfo = new TimerInfo();
        await function.Run(timerInfo, CancellationToken.None);
    }
}

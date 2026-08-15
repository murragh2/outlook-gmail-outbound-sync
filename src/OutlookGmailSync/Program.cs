using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Azure;
using OutlookGmailSync.Configuration;
using OutlookGmailSync.Services;
using Azure.Security.KeyVault.Secrets;
using Azure.Data.Tables;

var host = new HostBuilder()
    .ConfigureFunctionsWorkerDefaults()
    .ConfigureServices((context, services) =>
    {
        var config = context.Configuration;

        // Configuration
        services.Configure<AzureAdOptions>(config.GetSection("AzureAd"));
        services.Configure<SyncOptions>(config.GetSection("Sync"));

        // Add Azure clients
        services.AddSingleton(sp => 
            new DefaultAzureCredential(new DefaultAzureCredentialOptions 
            { 
                ExcludeEnvironmentCredential = true, 
                ExcludeInteractiveBrowserCredential = true 
            }));

        var storageUri = config.GetValue<string>("Storage:AccountUri");
        var keyVaultUri = config.GetValue<string>("KeyVault:VaultUri");

        services.AddSingleton(sp => 
        {
            var cred = sp.GetRequiredService<DefaultAzureCredential>();
            if (string.IsNullOrEmpty(storageUri))
                return new TableServiceClient(config.GetValue<string>("AzureWebJobsStorage"));
            return new TableServiceClient(new Uri(storageUri), cred);
        });

        services.AddSingleton(sp => 
        {
            var cred = sp.GetRequiredService<DefaultAzureCredential>();
            if (string.IsNullOrEmpty(keyVaultUri))
                throw new InvalidOperationException("KeyVault:VaultUri is not configured.");
            return new SecretClient(new Uri(keyVaultUri), cred);
        });

        services.AddHttpClient<IGraphTokenService, GraphTokenService>();

        services.AddSingleton<IGraphTokenService, GraphTokenService>();
        services.AddSingleton<IOutlookMailReader, OutlookMailReader>();
        services.AddSingleton<IMailForwarder, MailForwarder>();
        services.AddSingleton<IDeduplicationService, DeduplicationService>();
        services.AddSingleton<ISyncStateStore, SyncStateStore>();
    })
    .Build();

host.Run();

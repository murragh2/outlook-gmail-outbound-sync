using System.Text.Json;
using Azure.Identity;
using Azure.Security.KeyVault.Secrets;
using Microsoft.Identity.Client;

namespace OutlookGmailSync.AuthCli;

class Program
{
    static async Task Main(string[] args)
    {
        string? vaultUri = null;
        string? clientId = null;
        string tenantId = "consumers";
        bool saveLocal = false;

        for (int i = 0; i < args.Length; i++)
        {
            if (args[i] == "--vault-uri" && i + 1 < args.Length) vaultUri = args[++i];
            else if (args[i] == "--client-id" && i + 1 < args.Length) clientId = args[++i];
            else if (args[i] == "--tenant-id" && i + 1 < args.Length) tenantId = args[++i];
            else if (args[i] == "--save-local") saveLocal = true;
        }

        if (clientId == null)
        {
            Console.WriteLine("Usage: authcli --client-id <id> [--vault-uri <uri>] [--tenant-id consumers] [--save-local]");
            return;
        }

        var scopes = new[] { "Mail.Read", "Mail.Send", "offline_access" };

        var app = PublicClientApplicationBuilder.Create(clientId)
            .WithTenantId(tenantId)
            .WithRedirectUri("http://localhost:8400")
            .Build();

        Console.WriteLine("Acquiring token via device code...");

        AuthenticationResult result;
        try
        {
            result = await app.AcquireTokenWithDeviceCode(scopes, deviceCodeResult =>
            {
                Console.WriteLine(deviceCodeResult.Message);
                return Task.CompletedTask;
            }).ExecuteAsync();
        }
        catch (MsalException ex)
        {
            Console.WriteLine($"Error acquiring token: {ex.Message}");
            return;
        }

        Console.WriteLine("Successfully acquired token. Retrieving refresh token from cache...");

        byte[] cacheData = ((ITokenCacheSerializer)app.UserTokenCache).SerializeMsalV3();
        string cacheJson = System.Text.Encoding.UTF8.GetString(cacheData);

        var doc = JsonDocument.Parse(cacheJson);
        string? refreshToken = null;

        if (doc.RootElement.TryGetProperty("RefreshToken", out var rtElement))
        {
            foreach (var property in rtElement.EnumerateObject())
            {
                if (property.Value.TryGetProperty("secret", out var secretElement))
                {
                    refreshToken = secretElement.GetString();
                    break;
                }
            }
        }

        if (string.IsNullOrEmpty(refreshToken))
        {
            Console.WriteLine("Error: Could not extract refresh token from MSAL cache.");
            return;
        }

        if (saveLocal)
        {
            var localSettingsPath = Path.Combine(Directory.GetCurrentDirectory(), "src/OutlookGmailSync/local.settings.json");
            if (File.Exists(localSettingsPath))
            {
                var json = File.ReadAllText(localSettingsPath);
                var dict = JsonSerializer.Deserialize<Dictionary<string, object>>(json) ?? new();
                if (dict.TryGetValue("Values", out var valuesObj) && valuesObj is JsonElement valuesElement)
                {
                    var valuesDict = JsonSerializer.Deserialize<Dictionary<string, string>>(valuesElement.GetRawText()) ?? new();
                    valuesDict["AzureAd__GraphRefreshToken"] = refreshToken;
                    dict["Values"] = valuesDict;
                    File.WriteAllText(localSettingsPath, JsonSerializer.Serialize(dict, new JsonSerializerOptions { WriteIndented = true }));
                    Console.WriteLine("Success! Saved refresh token to local.settings.json under AzureAd__GraphRefreshToken.");
                }
            }
        }

        if (!string.IsNullOrEmpty(vaultUri))
        {
            Console.WriteLine("Storing refresh token in Key Vault as 'GraphRefreshToken'...");
            try
            {
                var cred = new DefaultAzureCredential();
                var secretClient = new SecretClient(new Uri(vaultUri), cred);
                await secretClient.SetSecretAsync("GraphRefreshToken", refreshToken);
                Console.WriteLine("Success! Refresh token stored in Key Vault.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Key Vault save skipped/error: {ex.Message}");
            }
        }
    }
}

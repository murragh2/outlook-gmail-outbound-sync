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

        for (int i = 0; i < args.Length; i++)
        {
            if (args[i] == "--vault-uri" && i + 1 < args.Length) vaultUri = args[++i];
            else if (args[i] == "--client-id" && i + 1 < args.Length) clientId = args[++i];
            else if (args[i] == "--tenant-id" && i + 1 < args.Length) tenantId = args[++i];
        }

        if (vaultUri == null || clientId == null)
        {
            Console.WriteLine("Usage: authcli --vault-uri <uri> --client-id <id> [--tenant-id consumers]");
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

        Console.WriteLine("Storing refresh token in Key Vault as 'GraphRefreshToken'...");

        var cred = new DefaultAzureCredential();
        var secretClient = new SecretClient(new Uri(vaultUri), cred);

        await secretClient.SetSecretAsync("GraphRefreshToken", refreshToken);

        Console.WriteLine("Success! Refresh token stored in Key Vault.");
    }
}

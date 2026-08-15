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
            Console.WriteLine("Saving short-lived Access Token (~1-hour TTL) securely to .NET User Secrets...");
            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = "dotnet",
                Arguments = $"user-secrets set \"AzureAd:GraphAccessToken\" \"{result.AccessToken}\" --project src/OutlookGmailSync",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false
            };
            var proc = System.Diagnostics.Process.Start(psi);
            proc?.WaitForExit();
            if (proc?.ExitCode == 0)
            {
                Console.WriteLine("Success! Saved short-lived Graph Access Token to .NET User Secrets store.");
                Console.WriteLine("Note: The refresh token was NOT saved. This access token will expire in 60 minutes.");
            }
            else
            {
                var err = proc?.StandardError.ReadToEnd();
                Console.WriteLine($"Warning: Failed to set user secrets: {err}");
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

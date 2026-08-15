using System.Net.Http.Json;
using Azure.Security.KeyVault.Secrets;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using OutlookGmailSync.Configuration;
using OutlookGmailSync.Models;

namespace OutlookGmailSync.Services;

public interface IGraphTokenService
{
    Task<string> GetAccessTokenAsync(CancellationToken ct);
}

public class GraphTokenService : IGraphTokenService
{
    private readonly HttpClient _httpClient;
    private readonly SecretClient _secretClient;
    private readonly AzureAdOptions _azureAdOptions;
    private readonly ILogger<GraphTokenService> _logger;

    public GraphTokenService(
        HttpClient httpClient,
        SecretClient secretClient,
        IOptions<AzureAdOptions> azureAdOptions,
        ILogger<GraphTokenService> logger)
    {
        _httpClient = httpClient;
        _secretClient = secretClient;
        _azureAdOptions = azureAdOptions.Value;
        _logger = logger;
    }

    public async Task<string> GetAccessTokenAsync(CancellationToken ct)
    {
        string refreshToken = _azureAdOptions.GraphRefreshToken;

        try
        {
            KeyVaultSecret secret = await _secretClient.GetSecretAsync("GraphRefreshToken", cancellationToken: ct);
            refreshToken = secret.Value;
        }
        catch (Azure.RequestFailedException ex) when (ex.Status == 403 || ex.Status == 401)
        {
            _logger.LogWarning("[LOCAL DEV] Key Vault access forbidden (403). Using local configuration token fallback.");
        }

        if (string.IsNullOrEmpty(refreshToken))
        {
            throw new InvalidOperationException("No refresh token available from Key Vault or local configuration.");
        }

        var tokenEndpoint = $"https://login.microsoftonline.com/{_azureAdOptions.TenantId}/oauth2/v2.0/token";
        
        var requestBody = new FormUrlEncodedContent(new[]
        {
            new KeyValuePair<string, string>("client_id", _azureAdOptions.ClientId),
            new KeyValuePair<string, string>("scope", "https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Mail.Send offline_access"),
            new KeyValuePair<string, string>("refresh_token", refreshToken),
            new KeyValuePair<string, string>("grant_type", "refresh_token")
        });

        if (!string.IsNullOrEmpty(_azureAdOptions.ClientSecret))
        {
            requestBody = new FormUrlEncodedContent(new[]
            {
                new KeyValuePair<string, string>("client_id", _azureAdOptions.ClientId),
                new KeyValuePair<string, string>("client_secret", _azureAdOptions.ClientSecret),
                new KeyValuePair<string, string>("scope", "https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Mail.Send offline_access"),
                new KeyValuePair<string, string>("refresh_token", refreshToken),
                new KeyValuePair<string, string>("grant_type", "refresh_token")
            });
        }

        var response = await _httpClient.PostAsync(tokenEndpoint, requestBody, ct);

        if (!response.IsSuccessStatusCode)
        {
            var errorContent = await response.Content.ReadAsStringAsync(ct);
            _logger.LogCritical("Failed to refresh token. Status Code: {StatusCode}, Error: {Error}", response.StatusCode, errorContent);
            throw new InvalidOperationException($"Failed to refresh token: {errorContent}");
        }

        var tokenResponse = await response.Content.ReadFromJsonAsync<TokenResponse>(cancellationToken: ct);

        if (tokenResponse == null || string.IsNullOrEmpty(tokenResponse.AccessToken))
        {
            throw new InvalidOperationException("Token response was null or missing access token.");
        }

        if (!string.IsNullOrEmpty(tokenResponse.RefreshToken) && tokenResponse.RefreshToken != refreshToken)
        {
            try
            {
                await _secretClient.SetSecretAsync("GraphRefreshToken", tokenResponse.RefreshToken, ct);
                _logger.LogInformation("Stored new refresh token in Key Vault.");
            }
            catch (Azure.RequestFailedException ex) when (ex.Status == 403 || ex.Status == 401)
            {
                _logger.LogWarning("[LOCAL DEV] Key Vault access forbidden (403). Skipped storing rotated refresh token in production Key Vault.");
            }
        }

        return tokenResponse.AccessToken;
    }
}

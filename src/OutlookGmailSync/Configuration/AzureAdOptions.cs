namespace OutlookGmailSync.Configuration;

public class AzureAdOptions
{
    public string TenantId { get; set; } = "consumers";
    public string ClientId { get; set; } = string.Empty;
    public string ClientSecret { get; set; } = string.Empty;
    public string GraphRefreshToken { get; set; } = string.Empty;
}

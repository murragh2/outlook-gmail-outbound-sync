using System.Security.Cryptography;
using System.Text;
using Azure;
using Azure.Data.Tables;

namespace OutlookGmailSync.Models;

public class ForwardedMailRecord : ITableEntity
{
    public string PartitionKey { get; set; } = string.Empty;
    public string RowKey { get; set; } = string.Empty;
    public DateTimeOffset? Timestamp { get; set; }
    public ETag ETag { get; set; }
    
    public string InternetMessageId { get; set; } = string.Empty;
    public string Subject { get; set; } = string.Empty;
    public DateTimeOffset SentDateTime { get; set; }
    public DateTimeOffset ForwardedAt { get; set; }

    public static ForwardedMailRecord Create(string internetMessageId, string subject, DateTimeOffset sentDateTime)
    {
        return new ForwardedMailRecord
        {
            PartitionKey = sentDateTime.ToString("yyyy-MM"),
            RowKey = ComputeHash(internetMessageId),
            InternetMessageId = internetMessageId,
            Subject = subject,
            SentDateTime = sentDateTime,
            ForwardedAt = DateTimeOffset.UtcNow
        };
    }

    public static string ComputeHash(string input)
    {
        var bytes = SHA256.HashData(Encoding.UTF8.GetBytes(input));
        return Convert.ToHexString(bytes).ToLowerInvariant();
    }
}

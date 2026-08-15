using Azure;
using Azure.Data.Tables;

namespace OutlookGmailSync.Models;

public class SyncState : ITableEntity
{
    public string PartitionKey { get; set; } = "SyncState";
    public string RowKey { get; set; } = "DeltaLink";
    public DateTimeOffset? Timestamp { get; set; }
    public ETag ETag { get; set; }

    public string? DeltaLink { get; set; }
    public DateTimeOffset LastSyncTime { get; set; }
    public bool IsInitialSyncComplete { get; set; }
}

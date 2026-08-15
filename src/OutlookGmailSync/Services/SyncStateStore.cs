using Azure.Data.Tables;
using OutlookGmailSync.Models;

namespace OutlookGmailSync.Services;

public interface ISyncStateStore
{
    Task<SyncState?> GetStateAsync(CancellationToken ct);
    Task SaveStateAsync(SyncState state, CancellationToken ct);
}

public class SyncStateStore : ISyncStateStore
{
    private readonly TableClient _tableClient;

    public SyncStateStore(TableServiceClient tableServiceClient)
    {
        _tableClient = tableServiceClient.GetTableClient("SyncState");
    }

    public async Task<SyncState?> GetStateAsync(CancellationToken ct)
    {
        await _tableClient.CreateIfNotExistsAsync(cancellationToken: ct);

        try
        {
            var response = await _tableClient.GetEntityAsync<SyncState>("SyncState", "DeltaLink", cancellationToken: ct);
            return response.Value;
        }
        catch (Azure.RequestFailedException ex) when (ex.Status == 404)
        {
            return null;
        }
    }

    public async Task SaveStateAsync(SyncState state, CancellationToken ct)
    {
        await _tableClient.CreateIfNotExistsAsync(cancellationToken: ct);
        await _tableClient.UpsertEntityAsync(state, TableUpdateMode.Replace, cancellationToken: ct);
    }
}

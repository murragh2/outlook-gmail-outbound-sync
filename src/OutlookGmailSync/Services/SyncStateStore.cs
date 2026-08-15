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

    private SyncState? _inMemoryState;

    public async Task<SyncState?> GetStateAsync(CancellationToken ct)
    {
        try
        {
            await _tableClient.CreateIfNotExistsAsync(cancellationToken: ct);
            var response = await _tableClient.GetEntityAsync<SyncState>("SyncState", "DeltaLink", cancellationToken: ct);
            return response.Value;
        }
        catch (Azure.RequestFailedException ex) when (ex.Status == 404)
        {
            return _inMemoryState;
        }
        catch (Azure.RequestFailedException ex) when (ex.Status == 403 || ex.Status == 401)
        {
            return _inMemoryState;
        }
    }

    public async Task SaveStateAsync(SyncState state, CancellationToken ct)
    {
        _inMemoryState = state;
        try
        {
            await _tableClient.CreateIfNotExistsAsync(cancellationToken: ct);
            await _tableClient.UpsertEntityAsync(state, TableUpdateMode.Replace, cancellationToken: ct);
        }
        catch (Azure.RequestFailedException ex) when (ex.Status == 403 || ex.Status == 401)
        {
            // In local dev without prod storage permissions, state persists in memory for local test session
        }
    }
}

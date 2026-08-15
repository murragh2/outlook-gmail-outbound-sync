using Azure.Data.Tables;
using OutlookGmailSync.Models;

namespace OutlookGmailSync.Services;

public interface IDeduplicationService
{
    Task<bool> HasBeenForwardedAsync(string internetMessageId, DateTimeOffset sentDateTime, CancellationToken ct);
    Task RecordForwardedAsync(string internetMessageId, string subject, DateTimeOffset sentDateTime, CancellationToken ct);
}

public class DeduplicationService : IDeduplicationService
{
    private readonly TableClient _tableClient;

    public DeduplicationService(TableServiceClient tableServiceClient)
    {
        _tableClient = tableServiceClient.GetTableClient("ForwardedMails");
    }

    public async Task<bool> HasBeenForwardedAsync(string internetMessageId, DateTimeOffset sentDateTime, CancellationToken ct)
    {
        await _tableClient.CreateIfNotExistsAsync(cancellationToken: ct);

        var partitionKey = sentDateTime.ToString("yyyy-MM");
        var rowKey = ForwardedMailRecord.ComputeHash(internetMessageId);

        try
        {
            var response = await _tableClient.GetEntityAsync<ForwardedMailRecord>(partitionKey, rowKey, cancellationToken: ct);
            return response.Value != null;
        }
        catch (Azure.RequestFailedException ex) when (ex.Status == 404)
        {
            return false;
        }
    }

    public async Task RecordForwardedAsync(string internetMessageId, string subject, DateTimeOffset sentDateTime, CancellationToken ct)
    {
        await _tableClient.CreateIfNotExistsAsync(cancellationToken: ct);

        var record = ForwardedMailRecord.Create(internetMessageId, subject, sentDateTime);
        await _tableClient.UpsertEntityAsync(record, TableUpdateMode.Replace, cancellationToken: ct);
    }
}

using FluentAssertions;
using OutlookGmailSync.Models;
using Xunit;

namespace OutlookGmailSync.Tests;

public class DeduplicationServiceTests
{
    [Fact]
    public void ForwardedMailRecord_Factory_CreatesCorrectRecord()
    {
        var id = "test-id";
        var subject = "test subject";
        var date = new DateTimeOffset(2026, 8, 14, 0, 0, 0, TimeSpan.Zero);

        var record = ForwardedMailRecord.Create(id, subject, date);

        record.PartitionKey.Should().Be("2026-08");
        record.InternetMessageId.Should().Be(id);
        record.Subject.Should().Be(subject);
        record.SentDateTime.Should().Be(date);
        record.ForwardedAt.Should().BeCloseTo(DateTimeOffset.UtcNow, TimeSpan.FromSeconds(5));
    }

    [Fact]
    public void ComputeHash_ReturnsSha256HexString()
    {
        var id = "test-id";
        var hash = ForwardedMailRecord.ComputeHash(id);

        hash.Should().NotBeNullOrEmpty();
        hash.Length.Should().Be(64);
    }
}

namespace OutlookGmailSync.Configuration;

public class SyncOptions
{
    public string GmailAddress { get; set; } = "murragh2@gmail.com";
    public string CronSchedule { get; set; } = "*/30 * * * * *";
}

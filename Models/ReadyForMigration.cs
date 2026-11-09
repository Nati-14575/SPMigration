namespace SPMigration.Models
{
    /// <summary>One SharePoint item that is eligible for migration.</summary>
    public class ReadyRecordItem
    {
        public int Id { get; set; }
        public string? Title { get; set; }
        public string? Subject { get; set; }
        public string? RecordType { get; set; }
        public string? ReferenceNumber { get; set; }
        public string? Source { get; set; }
        public string? Created { get; set; }
    }

    /// <summary>A list of ready items plus counts for one source.</summary>
    public class ReadyRecordsResult
    {
        public string Source { get; set; } = string.Empty;
        public int TotalFound { get; set; }
        public int AlreadyMigrated { get; set; }
        public int SkippedNotReady { get; set; }
        public int Count => Records.Count;
        public List<ReadyRecordItem> Records { get; set; } = new();
    }

    public class ReadyForMigrationResponse
    {
        public ReadyRecordsResult UserRecords { get; set; } = new();
        public ReadyRecordsResult IncomingRecords { get; set; } = new();
    }
}

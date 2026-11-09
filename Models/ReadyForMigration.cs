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
        /// <summary>Total ready items, even when Ids is truncated by the limit parameter.</summary>
        public int TotalReady { get; set; }
        /// <summary>Number of ids actually returned in this response.</summary>
        public int Count => Ids.Count;
        public List<int> Ids { get; set; } = new();
    }
}

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
        /// <summary>Ready ids (migrated ones only when includeMigrated=true). Equals failedIds + notMigratedIds otherwise.</summary>
        public List<int> Ids { get; set; } = new();
        /// <summary>Ids whose last logged attempt failed (in an *_errors.json log, never in a success log).</summary>
        public List<int> FailedIds { get; set; } = new();
        /// <summary>Why each failed id failed: its last logged error, classified. Same ids as FailedIds.</summary>
        public List<FailedRecordInfo> FailedDetails { get; set; } = new();
        /// <summary>Failed ids grouped by category (e.g. MissingRequiredField, NoFiles, NotFound).</summary>
        public Dictionary<string, List<int>> FailedByCategory { get; set; } = new();
        /// <summary>Number of failed ids per category (never truncated by limit).</summary>
        public Dictionary<string, int> FailureSummary { get; set; } = new();
        /// <summary>Ids never attempted: not in any success or error log.</summary>
        public List<int> NotMigratedIds { get; set; } = new();
    }

    /// <summary>The classified last failure of one id.</summary>
    public class FailedRecordInfo
    {
        public int Id { get; set; }
        /// <summary>"Migration" (payload step) or "FileInfo" (file-info update step).</summary>
        public string ErrorType { get; set; } = string.Empty;
        /// <summary>Best-effort classification from the logged error text.</summary>
        public string Category { get; set; } = string.Empty;
        /// <summary>The API's message when it was logged, otherwise the exception message.</summary>
        public string Reason { get; set; } = string.Empty;
        public DateTime FailedAt { get; set; }
    }
}

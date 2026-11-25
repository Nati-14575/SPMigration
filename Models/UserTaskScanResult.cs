namespace SPMigration.Models
{
    /// <summary>Records found through the tasks assigned to one user.</summary>
    public class UserTaskScanResult
    {
        public string User { get; set; } = string.Empty;
        /// <summary>"Internal" or "External".</summary>
        public string RunType { get; set; } = string.Empty;
        /// <summary>The task list that was read.</summary>
        public string Source { get; set; } = string.Empty;
        /// <summary>Record types asked for; empty when the list has no such column (External).</summary>
        public List<string> RecordTypes { get; set; } = new();
        public int TaskCount { get; set; }
        /// <summary>Distinct record ids (User_Records ids for Internal, Incoming folder ids for External), ascending.</summary>
        public List<int> Ids { get; set; } = new();
    }
}

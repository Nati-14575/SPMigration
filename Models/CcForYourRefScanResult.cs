namespace SPMigration.Models
{
    /// <summary>User_Records that have CC or ForYourRef tasks on the internal task list.</summary>
    public class CcForYourRefScanResult
    {
        /// <summary>The task list that was scanned.</summary>
        public string Source { get; set; } = string.Empty;
        /// <summary>Records found before the limit parameter is applied.</summary>
        public int TotalFound { get; set; }
        public int WithCc { get; set; }
        public int WithForYourRef { get; set; }
        public int WithBoth { get; set; }
        /// <summary>The User_Records ids that are (or, in a dry run, would be) migrated, ascending.</summary>
        public List<int> Ids { get; set; } = new();
    }
}

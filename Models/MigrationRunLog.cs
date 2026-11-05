namespace SPMigration.Models
{
    public class MigrationRunLog
    {
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public double DurationMinutes => (EndTime - StartTime).TotalMinutes;
        public int TotalProcessed { get; set; }
        public int TotalErrors { get; set; }

        // New properties
        public int? StartId { get; set; }
        public int? EndId { get; set; }
        public string Mode { get; set; } = string.Empty; // e.g. "Range" or "Resume"
    }
}

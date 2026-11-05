namespace SPMigration.Models
{
    public class MigrationSuccess
    {
        public int RecordId { get; set; }
        public string RecordsList { get; set; } = string.Empty;
        public string TasksList { get; set; } = string.Empty;
        public DateTime Timestamp { get; set; } = DateTime.UtcNow;
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public double DurationSeconds { get; set; }
    }

}

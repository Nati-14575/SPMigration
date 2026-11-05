namespace SPMigration.Models
{
    public class MigrationError
    {
        public int RecordId { get; set; }
        public string RecordsList { get; set; } = string.Empty;
        public string TasksList { get; set; } = string.Empty;
        public string InternalTasksList { get; set; } = string.Empty;
        public string InternalList { get; set; } = string.Empty;
        public string IncomingTasksList { get; set; } = string.Empty;
        public string IncomingList { get; set; } = string.Empty;
        public string ErrorMessage { get; set; } = string.Empty;
        public string? Payload { get; set; } // 👉 serialized payload JSON
        public DateTime Timestamp { get; set; } = DateTime.UtcNow;
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public double DurationSeconds { get; set; }
    }
}

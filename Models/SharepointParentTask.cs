namespace SPMigration.Models
{
    public class SharepointParentTask
    {
        public int ID { get; set; }
        public int Parent_x0020_ID { get; set; }
        public DateTime Created { get; set; }
        public DateTime Modified { get; set; }
        public string TaskStatus { get; set; } = string.Empty;
        public string Reassign_x0020_Remark { get; set; } = string.Empty;
        public string Author { get; set; } = string.Empty;
    }
}

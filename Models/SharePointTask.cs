namespace SPMigration.Models
{
    public class SharePointTask
    {
        public int ID { get; set; }
        public int Parent_x0020_ID { get; set; }
        public DateTime Created { get; set; }
        public DateTime Modified { get; set; }
        public string AssignedTo { get; set; } = string.Empty;
        public string TaskStatus { get; set; } = string.Empty;
        public string Reassign_x0020_Remark { get; set; } = string.Empty;
        public string Author { get; set; } = string.Empty;
        public int Parent_task_id { get; set; }


        public SharePointTask() { }

        public SharePointTask(
            int id,
            int parentId,
            DateTime created,
            DateTime modified,
            string assignedTo,
            string taskStatus,
            string reassignRemark,
            string author,
            string body,
            int parent_task_id)
        {
            ID = id;
            Parent_x0020_ID = parentId;
            Created = created;
            Modified = modified;
            AssignedTo = assignedTo ?? string.Empty;
            TaskStatus = taskStatus ?? string.Empty;
            Reassign_x0020_Remark = reassignRemark ?? string.Empty;
            Author = author ?? string.Empty;
            Parent_task_id = parent_task_id;
        }
    }
}

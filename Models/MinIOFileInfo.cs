namespace SPMigration.Models
{
    public class MinIOFileInfo
    {
        public string Bucket { get; set; }
        public string FileName { get; set; }
        public string ContentType { get; set; }
        public long Size { get; set; }
        public string Originalname { get; set; }
    }
}

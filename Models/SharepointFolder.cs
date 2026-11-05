namespace SPMigration.Models
{
    public class SharePointFolder
    {
        public int ID { get; set; }
        public string FileRef { get; set; }


        public string FileLeafRef { get; set; }
        
        public string Subject { get; set; }

        public string Sending_Organization_Name { get; set; }

        public string RecordFileName { get; set; }

        public string Incoming_Record_Date { get; set; }

        public string Reference_Number { get; set; }
        public string Delivery_Personnel_Name { get; set; }

        public string RecordNumber { get; set; }
        public string Author { get; set; }

    }

}
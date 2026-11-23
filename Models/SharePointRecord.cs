namespace SPMigration.Models
{
    public class SharePointRecord
    {
        public int ID { get; set; }
        public string Title { get; set; }
        public string FileLeafRef { get; set; }
        public int RecordId { get; set; }
        public string Reference_Number { get; set; }
        public string Original_Content { get; set; }
        public string Subject { get; set; }
        public DateTime? GregorianDate { get; set; }
        public string Recipient_Organization_Name { get; set; }
        public string CC_Organization_Name { get; set; }
        public string Prefix { get; set; }
        public string Suffix { get; set; }
        public string PrefixCC { get; set; }
        public string SuffixCC { get; set; }
        public string[] Receiving_Users { get; set; }
        public string[] Receiving_Department { get; set; }
        public string[] CC_To { get; set; }
        /// <summary>Second CC people/group column (Employee_CC); merged with CC_To into ccPosition.</summary>
        public string[] Employee_CC { get; set; } = Array.Empty<string>();
        /// <summary>"For your reference" people/group column (ForYouRefs); sent as forYourReferencePosition.</summary>
        public string[] ForYouRefs { get; set; } = Array.Empty<string>();
        public string RecordLanguage { get; set; }
        public string Record_Type { get; set; }
        public string Status { get; set; }
        public string Author { get; set; }
        public string Editor { get; set; }
        public DateTime Created { get; set; }
        public DateTime Modified { get; set; }
        public string Sending_Department { get; set; }
        public string Sending_Organization_Name { get; set; }
        public string DispatchedBy { get; set; }
        public DateTime DateOfDispatch { get; set; }
    }

}
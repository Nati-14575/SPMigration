
using SPMigration.Models;

public class IncomingRecordMigrationPayload
{
    public string referenceNumber { get; set; }
    public string id { get; set; }
    public string letterNumber { get; set; }
    public string unitKey { get; set; }
    public Metadata metadata { get; set; }
    public string createdAt { get; set; }
    public string comment { get; set; }
    public DateTime dispatchedDate { get; set; }
    public string dispatchedBy { get; set; }
    public string subject { get; set; }
    public string body { get; set; }
    public List<IncomingInstance> incomingInstances { get; set; }
}

public class Metadata
{
    public LocalizedText organizationName { get; set; }
    public LocalizedText deliveredBy { get; set; }
    public LocalizedText uploadedBy { get; set; }
}

public class LocalizedText
{
    public string en { get; set; }
    public string am { get; set; }
}

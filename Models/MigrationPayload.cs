using System.ComponentModel.DataAnnotations;

namespace SPMigration.Models
{
    public sealed class MigrationPayload
    {
        public int id { get; set; }
        public DateTime? createdAt { get; set; }
        public DateTime submittedDate { get; set; }
        public DateTime? completedDate { get; set; }
        public DateTime? dispatchedDate { get; set; }
        public string letterNumber { get; set; }
        public string referenceNumber { get; set; }
        public string initiatorId { get; set; }
        public string? dispatchedBy { get; set; }
        public string? completedBy { get; set; }
        public string preferredLanguage { get; set; }
        public string recordTypeKey { get; set; }
        public string statusKey { get; set; }
        public List<int> unitId { get; set; } = new();
        public string unitKey { get; set; }
        public MetadataBlock metadata { get; set; }
        public List<ContentItem> content { get; set; } = new();
        public List<IncomingInstance> incomingInstances { get; set; } = new();

    }

    public sealed class MetadataBlock
    {
        public string organizationName { get; set; }
        public string deliveredBy { get; set; }
    }

    public sealed class ContentItem
    {
        public string body { get; set; }
        public string subject { get; set; }
        public DateTime? date { get; set; }
        public List<string> nonSmartOfficeReceiverCCs { get; set; } = new();
        public List<string> nonSmartOfficeReceivers { get; set; } = new();
        public string sincerelyText { get; set; }
        public string prefix { get; set; }
        public string suffix { get; set; }
        public string prefixCC { get; set; }
        public string suffixCC { get; set; }
        public string[] receiverEmployees { get; set; }
        public string[]  receiverPositions { get; set; }
        public string[] ccPosition { get; set; }
        public List<OutgoingInstance> outgoingInstances { get; set; } = new();
        public List<object> outgoingHistories { get; set; } = new();
    }

    public sealed class OutgoingInstance
    {
        public string currentStepId { get; set; }
        public string dispatchedBy { get; set; }
        public DateTime? dispatchedAt { get; set; }
        public string taskStatus { get; set; }
        public string comment { get; set; }
    }


    public class FileInfoMigrationDto
    {
        [Required] // equivalent of @IsNotEmpty()
        public string Id { get; set; }

        // equivalent of attachments?: MinioFileOptions[]
        public List<MinIOFileInfo>? Attachments { get; set; }

        [Required] // equivalent of @IsNotEmpty()
        public MinIOFileInfo? FileInfo { get; set; }
    }
    public class IncomingInstance
    {
        public string status { get; set; } = string.Empty;

        public DateTime assignedAt { get; set; }

        public string id { get; set; }

        public string comment { get; set; } = string.Empty;

        public string? parentAssignmentId { get; set; }

        public string assignerUsername { get; set; }

        public string assigneeUsername { get; set; }

        public string assigneePositionKey { get; set; }
    }


}
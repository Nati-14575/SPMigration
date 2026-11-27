namespace SPMigration.Models
{
    public static class IamGroupSyncFailures
    {
        public const string InvalidName = "InvalidName";
        public const string InvalidPath = "InvalidPath";
        public const string DuplicateKey = "DuplicateKey";
        public const string PositionNotFound = "PositionNotFound";
        public const string PositionAmbiguous = "PositionAmbiguous";
        public const string PositionLookupError = "PositionLookupError";
        public const string CreateAuth = "CreateAuth";
        public const string CreateRejected = "CreateRejected";
        public const string CreateServerError = "CreateServerError";
        public const string CreateRequestError = "CreateRequestError";
    }

    /// <summary>One group of GroupsList.json: its position, the create call, and what happened to its members.</summary>
    public class IamGroupSyncResult
    {
        public string Group { get; set; } = string.Empty;
        public string Path { get; set; } = string.Empty;
        public string Key { get; set; } = string.Empty;
        /// <summary>The position name that was looked up and found (mapped name, or the last segment of the path).</summary>
        public string? PositionName { get; set; }
        /// <summary>"Mapping" or "PathName": where PositionName came from.</summary>
        public string? PositionSource { get; set; }
        public string? PositionId { get; set; }
        /// <summary>Created | Exists | Failed | Skipped | DryRun</summary>
        public string Status { get; set; } = string.Empty;
        public string? Category { get; set; }
        public string? Reason { get; set; }
        public int? HttpStatus { get; set; }
        public string? Response { get; set; }
        /// <summary>The member step for this group: Updated | Failed | Skipped | DryRun. Null when the group never got that far.</summary>
        public string? MemberStatus { get; set; }
        public int MembersSent { get; set; }
        public string? MemberReason { get; set; }
        public List<string> Warnings { get; set; } = new();
    }

    public class IamGroupSyncReport
    {
        public bool DryRun { get; set; }
        public int Total { get; set; }
        public int Created { get; set; }
        /// <summary>Groups IAM already had; they still get their members.</summary>
        public int Exists { get; set; }
        public int Failed { get; set; }
        public int Skipped { get; set; }
        public Dictionary<string, int> ProblemSummary { get; set; } = new();
        public List<IamGroupSyncResult> Results { get; set; } = new();
        /// <summary>The member step in full (its own report and retry file are written on a real run).</summary>
        public IamGroupMembersReport? Members { get; set; }
        public string? ReportFile { get; set; }
        /// <summary>Groups whose create call failed, in the GroupsList.json format, to post again.</summary>
        public string? RetryFile { get; set; }
        public string? FileError { get; set; }
    }
}

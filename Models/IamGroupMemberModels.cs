namespace SPMigration.Models
{
    // ---- Input: an existing IAM group (same shape as the IAM groups API returns; unknown fields are ignored) ----
    public class IamGroupInput
    {
        public string? Id { get; set; }
        public LocalizedName? Name { get; set; }
        public string? Key { get; set; }
        public string? UnitId { get; set; }
        /// <summary>Optional: the AD group to read members from when it is not named like Name.En.</summary>
        public string? AdGroupName { get; set; }
    }

    // ---- AD service (ADExplorer/GetGroupMembers) ----
    public class AdUser
    {
        /// <summary>"User" or "Group" when the AD service says so; absent in plain GetGroupMembers rows (treated as users).</summary>
        public string? Type { get; set; }
        public string? Name { get; set; }
        public string? SamAccountName { get; set; }
        public string? DisplayName { get; set; }
        public string? UserPrincipalName { get; set; }
        public string? EmailAddress { get; set; }
    }

    // ---- What is sent to the IAM API, one request per group ----
    /// <summary>The whole body of the request, nothing else.</summary>
    public class IamGroupMembersPayload
    {
        public string UnitKey { get; set; } = string.Empty;
        /// <summary>The IAM group's key.</summary>
        public string GroupKey { get; set; } = string.Empty;
        /// <summary>Usernames (SamAccountName) followed by the keys of nested groups.</summary>
        public List<string> Usernames { get; set; } = new();
    }

    public class IamGroupMembersRequest
    {
        public string Group { get; set; } = string.Empty;
        public string Method { get; set; } = string.Empty;
        /// <summary>IamGroups:GroupMembersUrl. Null in a dry run when it is not configured yet.</summary>
        public string? Url { get; set; }
        public IamGroupMembersPayload Body { get; set; } = new();
    }

    // ---- Report ----
    public static class IamGroupMemberStages
    {
        public const string Validation = "Validation";
        public const string AdLookup = "AdLookup";
        public const string Send = "Send";
    }

    public static class IamGroupMemberFailures
    {
        public const string InvalidInput = "InvalidInput";
        public const string DuplicateInput = "DuplicateInput";
        public const string AdGroupNotFound = "AdGroupNotFound";
        public const string AdHttpError = "AdHttpError";
        public const string AdTimeout = "AdTimeout";
        public const string AdBadResponse = "AdBadResponse";
        public const string AdNoMembers = "AdNoMembers";
        public const string IamAuth = "IamAuth";
        public const string IamRejected = "IamRejected";
        public const string IamServerError = "IamServerError";
        public const string IamRequestError = "IamRequestError";
    }

    public class IamGroupMembersResult
    {
        public string? GroupId { get; set; }
        public string Group { get; set; } = string.Empty;
        public string? Key { get; set; }
        public string? AdGroupName { get; set; }
        public string Status { get; set; } = string.Empty; // Updated | Failed | DryRun | Skipped
        /// <summary>Where it stopped: Validation, AdLookup or Send. Null when it went through.</summary>
        public string? Stage { get; set; }
        /// <summary>See IamGroupMemberFailures. Null when it went through.</summary>
        public string? Category { get; set; }
        public string? Reason { get; set; }
        /// <summary>Entries AD returned for the group.</summary>
        public int MembersFound { get; set; }
        /// <summary>Entries in usernames (users + nested groups).</summary>
        public int MembersSent { get; set; }
        /// <summary>How many of the usernames are nested groups. 0 for every group usually means GetGroupMembers does not return groups.</summary>
        public int NestedGroups { get; set; }
        public int? HttpStatus { get; set; }
        public string? Response { get; set; }
        public List<string> Warnings { get; set; } = new();
        public double DurationSeconds { get; set; }
    }

    public class IamGroupMembersReport
    {
        public bool DryRun { get; set; }
        public DateTime StartedAt { get; set; }
        public DateTime EndedAt { get; set; }
        public int Total { get; set; }
        /// <summary>Groups sent successfully (real run) or ready to send (dry run).</summary>
        public int Succeeded { get; set; }
        public int Failed { get; set; }
        public int Skipped { get; set; }
        /// <summary>Failed and skipped groups per category.</summary>
        public Dictionary<string, int> ProblemSummary { get; set; } = new();
        public List<IamGroupMembersResult> Results { get; set; } = new();
        /// <summary>The exact requests that will be (dry run) or were (real run) sent.</summary>
        public List<IamGroupMembersRequest> Requests { get; set; } = new();
        /// <summary>Real runs only: where the report and the retry file were written.</summary>
        public string? ReportFile { get; set; }
        public string? RetryFile { get; set; }
        public string? FileError { get; set; }
    }
}

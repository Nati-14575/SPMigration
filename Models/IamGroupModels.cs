using System.Text.Json;

namespace SPMigration.Models
{
    // ---- Input (shape of Updated.json) ----
    public class UpdatedPathEntry
    {
        public string Path { get; set; } = string.Empty;
        public List<UpdatedGroup> Groups { get; set; } = new();
    }

    public class UpdatedGroup
    {
        public string Name { get; set; } = string.Empty;
        public string Path { get; set; } = string.Empty;
        public string? Description { get; set; }
        public int Level { get; set; }
        public List<UpdatedMember> Members { get; set; } = new();
    }

    public class UpdatedMember
    {
        public string? Type { get; set; }
        public string? Name { get; set; }
        public string? SamAccountName { get; set; }
    }

    // ---- IAM payload ----
    public class LocalizedName
    {
        public string Am { get; set; } = string.Empty;
        public string En { get; set; } = string.Empty;
    }

    public class IamSubgroup
    {
        public LocalizedName Name { get; set; } = new();
        public string Key { get; set; } = string.Empty;
        public List<string> Members { get; set; } = new();
    }

    public class IamGroupPayload
    {
        public LocalizedName Name { get; set; } = new();
        public string Key { get; set; } = string.Empty;
        public string? UnitId { get; set; }
        public List<string> Members { get; set; } = new();
        public List<IamSubgroup> Subgroups { get; set; } = new();
    }

    // ---- Report ----
    public class IamGroupResult
    {
        public string Group { get; set; } = string.Empty;
        public string Key { get; set; } = string.Empty;
        public string Path { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty; // Created | Failed | DryRun | Skipped
        public int Members { get; set; }
        public int Subgroups { get; set; }
        public int? HttpStatus { get; set; }
        public string? Response { get; set; }
        public List<string> Warnings { get; set; } = new();
    }

    public class IamGroupMigrationReport
    {
        public bool DryRun { get; set; }
        public int Total { get; set; }
        public int Created { get; set; }
        public int Failed { get; set; }
        public int Skipped { get; set; }
        public List<IamGroupResult> Results { get; set; } = new();
        public List<IamGroupPayload>? Payloads { get; set; }
    }
}

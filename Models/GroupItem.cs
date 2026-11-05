using System.Text.Json;
using System.Text.Json.Serialization;

namespace SPMigration.Models
{
    /// <summary>Incoming group item (same shape as GroupsList.json).</summary>
    public class GroupItem
    {
        public string Name { get; set; } = string.Empty;
        public string Path { get; set; } = string.Empty;
        public string? Description { get; set; }
        public int Level { get; set; }
        public Guid Guid { get; set; }
    }

    public class GroupWithMembers : GroupItem
    {
        /// <summary>Raw members payload returned by the AD service (null if the lookup failed).</summary>
        public JsonElement? Members { get; set; }
        public string? Error { get; set; }
    }

    public class PathGroupsResult
    {
        public string Path { get; set; } = string.Empty;
        public List<GroupWithMembers> Groups { get; set; } = new();
    }
}

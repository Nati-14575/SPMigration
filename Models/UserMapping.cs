using System.Text.Json.Serialization;

namespace SPMigration.Models
{

    public class UserMapping
    {
        [JsonPropertyName("prevUserName")]
        public string PrevUserName { get; set; } = string.Empty;

        [JsonPropertyName("currentUserName")]
        public string CurrentUserName { get; set; } = string.Empty;
    }
}

using System.Text.Json;
using SPMigration.Models;

public class AdExplorerService
{
    private readonly HttpClient _httpClient;
    private readonly IConfiguration _configuration;
    private readonly ILogger<AdExplorerService> _logger;

    public AdExplorerService(HttpClient httpClient, IConfiguration configuration, ILogger<AdExplorerService> logger)
    {
        _httpClient = httpClient;
        _configuration = configuration;
        _logger = logger;
    }

    public async Task<JsonElement> GetAllGroupMembersAsync(string groupName, CancellationToken ct = default)
    {
        var baseUrl = _configuration["AdService:BaseUrl"]?.TrimEnd('/')
            ?? throw new InvalidOperationException("AdService:BaseUrl is not configured.");
        var url = $"{baseUrl}/ADExplorer/GetAllGroupMembers?group={Uri.EscapeDataString(groupName)}&recursive=true";

        using var response = await _httpClient.GetAsync(url, ct);
        response.EnsureSuccessStatusCode();
        using var doc = JsonDocument.Parse(await response.Content.ReadAsStringAsync(ct));
        return doc.RootElement.Clone();
    }

    /// <summary>Groups items by Path and attaches AD members to each group.</summary>
    public async Task<List<PathGroupsResult>> GetMembersByPathAsync(IEnumerable<GroupItem> items, CancellationToken ct = default)
    {
        var results = new List<PathGroupsResult>();
        using var gate = new SemaphoreSlim(5);

        foreach (var pathGroup in items.GroupBy(i => i.Path ?? string.Empty))
        {
            var tasks = pathGroup.Select(async g =>
            {
                var withMembers = new GroupWithMembers
                {
                    Name = g.Name, Path = g.Path, Description = g.Description, Level = g.Level, Guid = g.Guid
                };
                await gate.WaitAsync(ct);
                try
                {
                    withMembers.Members = await GetAllGroupMembersAsync(g.Name, ct);
                }
                catch (Exception ex) when (ex is not OperationCanceledException)
                {
                    _logger.LogWarning(ex, "Failed to get members for group {Group}", g.Name);
                    withMembers.Error = ex.Message;
                }
                finally { gate.Release(); }
                return withMembers;
            });

            results.Add(new PathGroupsResult
            {
                Path = pathGroup.Key,
                Groups = (await Task.WhenAll(tasks)).ToList()
            });
        }
        return results;
    }
}

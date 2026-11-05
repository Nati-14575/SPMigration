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

        // Skip groups whose name is the same as the last segment of their own path
        // (e.g. Name "X" under Path "A/B/X"); those are the unit's own group, not a sub-group.
        var candidates = items.Where(i => Normalize(i.Name) != Normalize(LastSegment(i.Path)));

        foreach (var pathGroup in candidates.GroupBy(i => i.Path ?? string.Empty))
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

    private static string LastSegment(string? path)
    {
        if (string.IsNullOrWhiteSpace(path)) return string.Empty;
        var idx = path.LastIndexOfAny(new[] { '/', '\\' });
        return idx >= 0 ? path[(idx + 1)..] : path;
    }

    // Trim and drop zero-width / non-breaking spaces so visually equal names compare equal.
    private static string Normalize(string? value) =>
        new string((value ?? string.Empty)
            .Where(c => c != '\u200B' && c != '\u200C' && c != '\u200D' && c != '\uFEFF')
            .ToArray())
            .Replace('\u00A0', ' ')
            .Trim()
            .ToLowerInvariant();
}

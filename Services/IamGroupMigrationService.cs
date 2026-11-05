using System.Text.Json;
using System.Text.RegularExpressions;
using SPMigration.Models;

public class IamGroupMigrationService
{
    private readonly ApiService _api;
    private readonly IConfiguration _config;
    private readonly ILogger<IamGroupMigrationService> _logger;

    public IamGroupMigrationService(ApiService api, IConfiguration config,
        ILogger<IamGroupMigrationService> logger)
    {
        _api = api; _config = config; _logger = logger;
    }

    public async Task<IamGroupMigrationReport> RunAsync(List<UpdatedPathEntry> input, string? unitId, bool dryRun, CancellationToken ct = default)
    {
        var groups = input.SelectMany(p => p.Groups).ToList();
        var byName = groups.GroupBy(g => Norm(g.Name)).ToDictionary(g => g.Key, g => g.First());
        var createUrl = _config["IamGroups:CreateGroupUrl"];

        if (!dryRun && string.IsNullOrWhiteSpace(unitId))
            throw new ArgumentException("unitId is required when dryRun=false.");
        if (!dryRun && string.IsNullOrWhiteSpace(createUrl))
            throw new InvalidOperationException("IamGroups:CreateGroupUrl is not configured.");

        var report = new IamGroupMigrationReport { DryRun = dryRun, Total = groups.Count };
        if (dryRun) report.Payloads = new();

        foreach (var g in groups)
        {
            ct.ThrowIfCancellationRequested();
            var result = new IamGroupResult { Group = g.Name, Path = g.Path, Key = ToKey(g.Name) };
            report.Results.Add(result);

            try
            {
                var payload = new IamGroupPayload
                {
                    Name = ToName(g.Name, g.Description),
                    Key = result.Key,
                    UnitId = string.IsNullOrWhiteSpace(unitId) ? null : unitId.Trim(),
                    Members = Users(g.Members)
                };

                // Subgroups: members of Type "Group" (one level).
                var usedKeys = new HashSet<string> { payload.Key };
                foreach (var sg in g.Members.Where(m => m.Type == "Group" && !string.IsNullOrWhiteSpace(m.Name)))
                {
                    var key = ToKey(sg.Name!);
                    if (!usedKeys.Add(key)) continue;

                    // Members come only from Updated.json; a subgroup not listed there has none.
                    var members = byName.TryGetValue(Norm(sg.Name!), out var known)
                        ? Users(known.Members)
                        : new List<string>();

                    payload.Subgroups.Add(new IamSubgroup { Name = ToName(sg.Name!, null), Key = key, Members = members });
                }

                result.Members = payload.Members.Count;
                result.Subgroups = payload.Subgroups.Count;

                if (dryRun)
                {
                    result.Status = "DryRun";
                    report.Payloads!.Add(payload);
                    continue;
                }

                var (ok, status, body) = await _api.PostJsonAsync(createUrl!, payload);
                result.HttpStatus = status;
                result.Response = body.Length > 2000 ? body[..2000] : body; // may list unknown usernames
                result.Status = ok ? "Created" : "Failed";
                if (ok) report.Created++; else report.Failed++;
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                _logger.LogError(ex, "Failed to process group {Group}", g.Name);
                result.Status = "Failed";
                result.Response = ex.Message;
                report.Failed++;
            }
        }
        return report;
    }

    private static List<string> Users(IEnumerable<UpdatedMember> members) =>
        members.Where(m => m.Type == "User" && !string.IsNullOrWhiteSpace(m.SamAccountName))
               .Select(m => m.SamAccountName!.Trim())
               .Distinct(StringComparer.OrdinalIgnoreCase).ToList();

    // Amharic description when present, otherwise fall back to the English name.
    private static LocalizedName ToName(string name, string? description) => new()
    {
        En = name.Trim(),
        Am = !string.IsNullOrWhiteSpace(description) && Regex.IsMatch(description, @"[ሀ-፿]")
            ? description.Trim() : name.Trim()
    };

    private static string ToKey(string name) =>
        Regex.Replace(name.ToLowerInvariant(), @"[^a-z0-9]+", "-").Trim('-');

    private static string Norm(string v) => v.Trim().ToLowerInvariant();
}

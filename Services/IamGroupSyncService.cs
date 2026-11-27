using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using SPMigration.Models;

/// <summary>
/// For each group of GroupsList.json: finds its IAM position (mapping file first, then the last segment of its path), creates the group
/// in IAM, then adds the AD members of every group that was created or already existed (IamGroupMembersService). Nothing aborts the run.
/// </summary>
public class IamGroupSyncService
{
    private readonly ApiService _api;
    private readonly IamPositionResolver _positions;
    private readonly IamGroupMembersService _members;
    private readonly IConfiguration _config;
    private readonly ILogger<IamGroupSyncService> _logger;

    public IamGroupSyncService(ApiService api, IamPositionResolver positions, IamGroupMembersService members,
        IConfiguration config, ILogger<IamGroupSyncService> logger)
    {
        _api = api; _positions = positions; _members = members; _config = config; _logger = logger;
    }

    public async Task<IamGroupSyncReport> RunAsync(List<GroupItem> groups, List<UserMapping>? mappings, string? unitId, bool dryRun, CancellationToken ct = default)
    {
        unitId = !string.IsNullOrWhiteSpace(unitId) ? unitId.Trim() : _config["IamGroups:UnitId"]?.Trim();
        var createUrl = _config["IamGroups:CreateGroupUrl"]?.Trim();
        if (!dryRun)
        {
            if (string.IsNullOrEmpty(unitId)) throw new InvalidOperationException("IamGroups:UnitId is not configured in appsettings.json (required when dryRun=false).");
            if (string.IsNullOrEmpty(createUrl)) throw new InvalidOperationException("IamGroups:CreateGroupUrl is not configured.");
            if (!_positions.IsConfigured) throw new InvalidOperationException("IamGroups:PositionLookupUrl is not configured.");
        }

        // old name -> IAM position name. Entries without a currentUserName (the "ambiguous" ones) cannot be used.
        var mapped = new Dictionary<string, string>();
        foreach (var m in mappings ?? new List<UserMapping>())
        {
            var prev = Norm(m.PrevUserName);
            var current = Clean(m.CurrentUserName);
            if (prev.Length > 0 && current.Length > 0) mapped.TryAdd(prev, current);
        }

        var report = new IamGroupSyncReport { DryRun = dryRun, Total = groups.Count };
        var usedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var inputs = new List<IamGroupInput>();
        var byKey = new Dictionary<string, IamGroupSyncResult>(StringComparer.OrdinalIgnoreCase);
        var failedGroups = new List<GroupItem>();

        foreach (var g in groups)
        {
            ct.ThrowIfCancellationRequested();
            var name = Clean(g.Name);
            var leaf = Clean(LastSegment(g.Path));
            var result = new IamGroupSyncResult { Group = name, Path = g.Path ?? string.Empty, Key = IamGroupMigrationService.ToKey(name) };
            report.Results.Add(result);

            void Stop(string status, string category, string reason)
            {
                result.Status = status; result.Category = category; result.Reason = reason;
            }

            if (result.Key.Length == 0) { Stop("Skipped", IamGroupSyncFailures.InvalidName, "The group name gives an empty key."); continue; }
            if (leaf is "" or "." or "..") { Stop("Skipped", IamGroupSyncFailures.InvalidPath, $"The path '{g.Path}' has no usable last segment."); continue; }
            if (!usedKeys.Add(result.Key)) { Stop("Skipped", IamGroupSyncFailures.DuplicateKey, "Another group in the list has the same key."); continue; }

            // 1) Position: mapped name first, then the last segment of the path itself.
            var candidates = new List<(string Name, string Source)>();
            if (mapped.TryGetValue(Norm(leaf), out var mappedName)) candidates.Add((mappedName, "Mapping"));
            if (!candidates.Any(c => c.Name.Equals(leaf, StringComparison.OrdinalIgnoreCase))) candidates.Add((leaf, "PathName"));

            if (!_positions.IsConfigured)
            {
                // Dry run only (a real run needs the lookup): everything else can still be checked.
                result.Warnings.Add("IamGroups:PositionLookupUrl is not configured: the position was not looked up.");
                result.PositionName = candidates[0].Name; result.PositionSource = candidates[0].Source;
            }
            else
            {
                PositionLookupResult? last = null;
                var found = false;
                foreach (var (candidate, source) in candidates)
                {
                    last = await _positions.LookupAsync(candidate);
                    if (last.Trace != null) result.Calls.Add(last.Trace);
                    if (last.Outcome == PositionLookupOutcome.Found)
                    {
                        result.PositionName = candidate; result.PositionSource = source; result.PositionId = last.Id;
                        found = true;
                        break;
                    }
                    if (last.Outcome != PositionLookupOutcome.NotFound) break; // Ambiguous / Error: do not guess with another name
                }

                if (!found)
                {
                    var tried = string.Join(", ", candidates.Select(c => $"'{c.Name}'"));
                    switch (last!.Outcome)
                    {
                        case PositionLookupOutcome.Ambiguous:
                            Stop("Skipped", IamGroupSyncFailures.PositionAmbiguous, $"{last.Detail} Tried {tried}."); break;
                        case PositionLookupOutcome.Error:
                            Stop("Failed", IamGroupSyncFailures.PositionLookupError, last.Detail ?? "Position lookup failed.");
                            result.HttpStatus = last.HttpStatus; failedGroups.Add(g); break;
                        default:
                            Stop("Skipped", IamGroupSyncFailures.PositionNotFound, $"No IAM position found. Tried {tried}."); break;
                    }
                    continue;
                }
                if (result.PositionSource == "PathName" && mapped.ContainsKey(Norm(leaf)))
                    result.Warnings.Add("The mapped position name was not found; the path name was used instead.");
            }

            // 2) Create the group.
            if (dryRun)
            {
                result.Status = "DryRun";
            }
            else
            {
                await CreateAsync(result, g, name, unitId!, createUrl!, failedGroups);
            }

            if (result.Status is "Created" or "Exists" or "DryRun")
            {
                inputs.Add(new IamGroupInput
                {
                    Name = IamGroupMigrationService.ToName(name, g.Description),
                    Key = result.Key,
                    // AD knows the group by its original name (which may hold invisible characters the key and En name drop).
                    AdGroupName = g.Name
                });
                byKey[result.Key] = result;
            }
        }

        // 3) Members for every group that now exists: users first, nested group keys after (all groups were created above).
        if (inputs.Count > 0)
        {
            report.Members = await _members.RunAsync(inputs, dryRun, ct);
            foreach (var mr in report.Members.Results)
            {
                if (mr.Key == null || !byKey.TryGetValue(mr.Key, out var r)) continue;
                r.MemberStatus = mr.Status; r.MembersSent = mr.MembersSent; r.MemberReason = mr.Reason;
                r.Warnings.AddRange(mr.Warnings);
            }
        }

        foreach (var r in report.Results)
        {
            switch (r.Status)
            {
                case "Created": report.Created++; break;
                case "Exists": report.Exists++; break;
                case "Skipped": report.Skipped++; break;
                case "Failed": report.Failed++; break;
                case "DryRun": report.Ready++; break;
            }
            if (r.Category != null) report.ProblemSummary[r.Category] = report.ProblemSummary.GetValueOrDefault(r.Category) + 1;
        }

        if (!dryRun) WriteFiles(report, failedGroups);
        return report;
    }

    // IAM's message for a group that is already there: '972b42fd-..., ministry-of-agriculture' is already in use for Unit Id, Key.
    private static readonly System.Text.RegularExpressions.Regex AlreadyInUse =
        new(@"'(?<unit>[^',]*),\s*(?<key>[^']+)'\s+is already in use", System.Text.RegularExpressions.RegexOptions.IgnoreCase);

    private async Task CreateAsync(IamGroupSyncResult result, GroupItem g, string name, string unitId, string createUrl, List<GroupItem> failedGroups)
    {
        try
        {
            // A dictionary so that parentGroupId is sent as an explicit null.
            var body = new Dictionary<string, object?>
            {
                ["name"] = IamGroupMigrationService.ToName(name, g.Description),
                ["key"] = result.Key,
                ["unitId"] = unitId,
                ["positionId"] = result.PositionId,
                ["parentGroupId"] = null
            };
            var bodyJson = ApiService.ToJson(body);
            var (ok, status, text) = await _api.PostJsonAsync(createUrl, body);
            result.HttpStatus = status;
            result.Response = text.Length > 2000 ? text[..2000] : text;
            result.Calls.Add(new IamCallTrace("CreateGroup", "POST", createUrl, bodyJson, status, result.Response));
            _logger.LogInformation("[iam] POST {Url} body={Body} -> {Status}", createUrl, bodyJson, status);

            if (ok) { result.Status = "Created"; return; }

            // "Already exists" is not a failure: the group is there, so it still gets its members. IAM answers 400 with
            // "'<unitId>, <key>' is already in use for Unit Id, Key."; the key in that message is the one the members are added to.
            var inUse = AlreadyInUse.Match(text);
            if (status == 409 || (status is 400 or 422 &&
                    (inUse.Success ||
                     text.Contains("already exist", StringComparison.OrdinalIgnoreCase) ||
                     text.Contains("duplicate", StringComparison.OrdinalIgnoreCase))))
            {
                result.Status = "Exists";
                result.Reason = "IAM already has a group with this key.";
                var keyFromIam = inUse.Success ? inUse.Groups["key"].Value.Trim() : string.Empty;
                if (keyFromIam.Length > 0 && !keyFromIam.Equals(result.Key, StringComparison.OrdinalIgnoreCase))
                {
                    result.Warnings.Add($"IAM reports the key '{keyFromIam}' (sent '{result.Key}'); the members are added to '{keyFromIam}'.");
                    result.Key = keyFromIam;
                }
                return;
            }

            result.Status = "Failed";
            result.Category = status is 401 or 403 ? IamGroupSyncFailures.CreateAuth
                : status >= 500 ? IamGroupSyncFailures.CreateServerError
                : IamGroupSyncFailures.CreateRejected;
            result.Reason = $"IAM answered {status}.";
            failedGroups.Add(g);
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            _logger.LogError(ex, "Failed to create IAM group {Group}", name);
            result.Status = "Failed";
            result.Category = IamGroupSyncFailures.CreateRequestError;
            result.Reason = ex.Message;
            result.Calls.Add(new IamCallTrace("CreateGroup", "POST", createUrl, null, null, ex.Message));
            failedGroups.Add(g);
        }
    }

    private void WriteFiles(IamGroupSyncReport report, List<GroupItem> failedGroups)
    {
        try
        {
            var dir = _config["IamGroups:ReportPath"];
            if (string.IsNullOrWhiteSpace(dir)) dir = Path.Combine(AppContext.BaseDirectory, "iam-group-reports");
            Directory.CreateDirectory(dir);

            var stamp = DateTime.UtcNow.ToString("yyyyMMdd-HHmmss");
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
                Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            };

            report.ReportFile = Path.Combine(dir, $"iam-groups-sync_{stamp}_report.json");
            File.WriteAllText(report.ReportFile, JsonSerializer.Serialize(report, options));

            if (failedGroups.Count > 0)
            {
                report.RetryFile = Path.Combine(dir, $"iam-groups-sync_{stamp}_retry.json");
                File.WriteAllText(report.RetryFile, JsonSerializer.Serialize(failedGroups, options));
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Could not write the IAM group sync report files");
            report.FileError = ex.Message;
        }
    }

    private static string LastSegment(string? path)
    {
        if (string.IsNullOrWhiteSpace(path)) return string.Empty;
        var idx = path.LastIndexOfAny(new[] { '/', '\\' });
        return idx >= 0 ? path[(idx + 1)..] : path;
    }

    // Drops zero-width characters, turns non-breaking spaces into spaces, collapses runs of spaces and trims.
    private static string Clean(string? value)
    {
        var chars = (value ?? string.Empty).Where(c => c != '​' && c != '‌' && c != '‍' && c != '﻿')
            .Select(c => c == ' ' ? ' ' : c).ToArray();
        return System.Text.RegularExpressions.Regex.Replace(new string(chars), @"\s+", " ").Trim();
    }

    private static string Norm(string? value) => Clean(value).ToLowerInvariant();
}

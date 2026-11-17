using System.Diagnostics;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using SPMigration.Models;

/// <summary>
/// For each existing IAM group: reads the group's members from the AD service and sends the member usernames
/// (SamAccountName) to the IAM group-update API. Every group ends with a result; nothing aborts the run.
/// </summary>
public class IamGroupMembersService
{
    private const int MaxParallelAdLookups = 5;
    private const int MaxResponseChars = 2000;

    private readonly AdExplorerService _ad;
    private readonly ApiService _api;
    private readonly IConfiguration _config;
    private readonly ILogger<IamGroupMembersService> _logger;

    public IamGroupMembersService(AdExplorerService ad, ApiService api, IConfiguration config,
        ILogger<IamGroupMembersService> logger)
    {
        _ad = ad; _api = api; _config = config; _logger = logger;
    }

    // One group after its AD lookup: the result so far, plus the request when it can be sent.
    private sealed record Prepared(IamGroupInput Input, IamGroupMembersResult Result, IamGroupMembersRequest? Request);

    public async Task<IamGroupMembersReport> RunAsync(List<IamGroupInput> input, string unitKey, bool dryRun, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(unitKey))
            throw new ArgumentException("unitKey is required.", nameof(unitKey));
        unitKey = unitKey.Trim();

        // The URL is only needed to send; a dry run works without it so the payloads can be checked first.
        var url = _config["IamGroups:GroupMembersUrl"]?.Trim();
        var method = new HttpMethod((_config["IamGroups:GroupMembersMethod"] ?? "POST").Trim().ToUpperInvariant());
        if (!dryRun && string.IsNullOrWhiteSpace(url))
            throw new InvalidOperationException("IamGroups:GroupMembersUrl is not configured.");

        var report = new IamGroupMembersReport { DryRun = dryRun, StartedAt = DateTime.UtcNow, Total = input.Count };

        // 1. Validate and read the AD members, a few groups at a time.
        var seenGroups = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using var gate = new SemaphoreSlim(MaxParallelAdLookups);
        var prepared = await Task.WhenAll(input.Select(g => PrepareAsync(g, unitKey, url, method, seenGroups, gate, ct)));

        // 2. Send one group after another (the IAM token is shared).
        foreach (var p in prepared)
        {
            ct.ThrowIfCancellationRequested();
            if (p.Request != null)
            {
                if (dryRun) p.Result.Status = "DryRun";
                else await SendAsync(p, method, ct);
            }
            Tally(report, p);
        }

        report.EndedAt = DateTime.UtcNow;
        if (!dryRun) WriteFiles(report, prepared);
        return report;
    }

    private async Task<Prepared> PrepareAsync(IamGroupInput g, string unitKey, string? url, HttpMethod method,
        HashSet<string> seenGroups, SemaphoreSlim gate, CancellationToken ct)
    {
        var sw = Stopwatch.StartNew();
        var groupKey = g.Key?.Trim() ?? string.Empty;
        var result = new IamGroupMembersResult
        {
            GroupId = g.Id?.Trim(),
            Group = g.Name?.En?.Trim() ?? string.Empty,
            Key = groupKey
        };

        Prepared Stop(string status, string stage, string category, string reason)
        {
            result.Status = status; result.Stage = stage; result.Category = category; result.Reason = reason;
            result.DurationSeconds = sw.Elapsed.TotalSeconds;
            return new Prepared(g, result, null);
        }

        // The group's key is sent as groupName, so it is always required.
        if (string.IsNullOrEmpty(groupKey))
            return Stop("Failed", IamGroupMemberStages.Validation, IamGroupMemberFailures.InvalidInput, "key is missing (it is sent as groupName).");

        // The members are read from AD by name: name.en, unless adGroupName says otherwise.
        var adName = (!string.IsNullOrWhiteSpace(g.AdGroupName) ? g.AdGroupName : g.Name?.En)?.Trim();
        if (string.IsNullOrEmpty(adName))
            return Stop("Failed", IamGroupMemberStages.Validation, IamGroupMemberFailures.InvalidInput, "name.en is missing (and no adGroupName was given).");
        result.AdGroupName = adName;
        if (string.IsNullOrEmpty(result.Group)) result.Group = adName;

        // The API only receives unitKey + groupName, so a repeated key would send the same payload twice.
        lock (seenGroups)
        {
            if (!seenGroups.Add(groupKey))
                return Stop("Skipped", IamGroupMemberStages.Validation, IamGroupMemberFailures.DuplicateInput, "The same group key appears more than once in the input.");
        }

        List<AdUser> users;
        await gate.WaitAsync(ct);
        try
        {
            users = await _ad.GetGroupMembersAsync(adName, ct);
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested) { throw; }
        catch (OperationCanceledException ex)
        {
            return Stop("Failed", IamGroupMemberStages.AdLookup, IamGroupMemberFailures.AdTimeout, ex.Message);
        }
        catch (HttpRequestException ex)
        {
            var category = ex.StatusCode == System.Net.HttpStatusCode.NotFound
                ? IamGroupMemberFailures.AdGroupNotFound : IamGroupMemberFailures.AdHttpError;
            _logger.LogWarning(ex, "AD lookup failed for group {Group}", adName);
            return Stop("Failed", IamGroupMemberStages.AdLookup, category, ex.Message);
        }
        catch (JsonException ex)
        {
            return Stop("Failed", IamGroupMemberStages.AdLookup, IamGroupMemberFailures.AdBadResponse, ex.Message);
        }
        finally { gate.Release(); }

        // userNames = the users' SamAccountName, then the keys of nested groups (same key rule as the group creation run).
        var userNames = new List<string>();
        var seenNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var m in users)
        {
            if (string.Equals(m.Type, "Group", StringComparison.OrdinalIgnoreCase))
            {
                var nestedName = new[] { m.Name, m.DisplayName, m.SamAccountName }.FirstOrDefault(n => !string.IsNullOrWhiteSpace(n));
                var nestedKey = nestedName == null ? string.Empty : IamGroupMigrationService.ToKey(nestedName);
                if (string.IsNullOrEmpty(nestedKey))
                    result.Warnings.Add("A nested group in AD has no name and was left out.");
                else if (nestedKey.Equals(groupKey, StringComparison.OrdinalIgnoreCase))
                    result.Warnings.Add($"Nested group '{nestedName}' has the same key as the group itself and was left out.");
                else if (seenNames.Add(nestedKey))
                {
                    userNames.Add(nestedKey);
                    result.NestedGroups++;
                }
                continue;
            }

            var sam = m.SamAccountName?.Trim();
            if (string.IsNullOrEmpty(sam))
                result.Warnings.Add($"AD member '{m.DisplayName}' has no SamAccountName and was left out.");
            else if (seenNames.Add(sam))
                userNames.Add(sam);
        }
        result.MembersFound = users.Count;

        // Never send an empty list: if the API replaces members, that would wipe the group.
        if (userNames.Count == 0)
            return Stop("Skipped", IamGroupMemberStages.AdLookup, IamGroupMemberFailures.AdNoMembers, "AD returned no users or groups for this group; nothing was sent.");

        result.MembersSent = userNames.Count;
        result.DurationSeconds = sw.Elapsed.TotalSeconds;
        var request = new IamGroupMembersRequest
        {
            Group = result.Group,
            Method = method.Method,
            Url = url,
            Body = new IamGroupMembersPayload { UnitKey = unitKey, GroupName = groupKey, UserNames = userNames }
        };
        return new Prepared(g, result, request);
    }

    private async Task SendAsync(Prepared p, HttpMethod method, CancellationToken ct)
    {
        var sw = Stopwatch.StartNew();
        var result = p.Result;
        try
        {
            var (ok, status, body) = await _api.SendJsonAsync(method, p.Request!.Url!, p.Request.Body);
            result.HttpStatus = status;
            result.Response = body.Length > MaxResponseChars ? body[..MaxResponseChars] : body; // may name unknown usernames
            if (ok)
            {
                result.Status = "Updated";
                return;
            }

            result.Status = "Failed";
            result.Stage = IamGroupMemberStages.Send;
            result.Category = status is 401 or 403 ? IamGroupMemberFailures.IamAuth
                : status >= 500 ? IamGroupMemberFailures.IamServerError
                : IamGroupMemberFailures.IamRejected;
            result.Reason = $"IAM answered {status}.";
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested) { throw; }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to update IAM group {Group}", result.Group);
            result.Status = "Failed";
            result.Stage = IamGroupMemberStages.Send;
            result.Category = IamGroupMemberFailures.IamRequestError;
            result.Reason = ex.Message;
        }
        finally { result.DurationSeconds += sw.Elapsed.TotalSeconds; }
    }

    private static void Tally(IamGroupMembersReport report, Prepared p)
    {
        report.Results.Add(p.Result);
        if (p.Request != null) report.Requests.Add(p.Request);

        switch (p.Result.Status)
        {
            case "Updated" or "DryRun": report.Succeeded++; break;
            case "Skipped": report.Skipped++; break;
            default: report.Failed++; break;
        }
        if (p.Result.Category != null)
            report.ProblemSummary[p.Result.Category] = report.ProblemSummary.GetValueOrDefault(p.Result.Category) + 1;
    }

    // Real runs: the full report, plus the failed groups in the input format so they can be posted again.
    private void WriteFiles(IamGroupMembersReport report, IEnumerable<Prepared> prepared)
    {
        try
        {
            var dir = _config["IamGroups:ReportPath"];
            if (string.IsNullOrWhiteSpace(dir)) dir = Path.Combine(AppContext.BaseDirectory, "iam-group-reports");
            Directory.CreateDirectory(dir);

            var stamp = report.StartedAt.ToString("yyyyMMdd-HHmmss");
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
                Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            };

            var reportFile = Path.Combine(dir, $"iam-group-members_{stamp}_report.json");
            File.WriteAllText(reportFile, JsonSerializer.Serialize(report, options));
            report.ReportFile = reportFile;

            var retry = prepared.Where(p => p.Result.Status == "Failed").Select(p => p.Input).ToList();
            if (retry.Count > 0)
            {
                var retryFile = Path.Combine(dir, $"iam-group-members_{stamp}_retry.json");
                File.WriteAllText(retryFile, JsonSerializer.Serialize(retry, options));
                report.RetryFile = retryFile;
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Could not write the IAM group members report files");
            report.FileError = ex.Message;
        }
    }
}

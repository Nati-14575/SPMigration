using System.Text.Json;

public enum PositionLookupOutcome { Found, NotFound, Ambiguous, Error }

public record PositionLookupResult(PositionLookupOutcome Outcome, string? Id, string? Detail, int? HttpStatus = null);

/// <summary>
/// Asks the IAM position lookup API for the id of a position name:
/// POST IamGroups:PositionLookupUrl { "names": [ "<name>" ] } (one name per call) answered with
/// [ { "query": "<name>", "matches": [ { "id", "name": { "am", "en" }, "key" } ] } ].
/// No match = NotFound. One match = Found. Several matches = the one whose name is exactly the query, otherwise Ambiguous.
/// Answers are cached by name for the life of this instance.
/// </summary>
public class IamPositionResolver
{
    private readonly ApiService _api;
    private readonly IConfiguration _config;
    private readonly ILogger<IamPositionResolver> _logger;
    private readonly Dictionary<string, PositionLookupResult> _cache = new(StringComparer.OrdinalIgnoreCase);

    public IamPositionResolver(ApiService api, IConfiguration config, ILogger<IamPositionResolver> logger)
    {
        _api = api; _config = config; _logger = logger;
    }

    public bool IsConfigured => !string.IsNullOrWhiteSpace(_config["IamGroups:PositionLookupUrl"]);

    public async Task<PositionLookupResult> LookupAsync(string name)
    {
        if (_cache.TryGetValue(name, out var cached)) return cached;
        var result = await LookupCoreAsync(name);
        if (result.Outcome != PositionLookupOutcome.Error) _cache[name] = result; // a failed call may work next time
        return result;
    }

    private async Task<PositionLookupResult> LookupCoreAsync(string name)
    {
        var url = _config["IamGroups:PositionLookupUrl"]?.Trim();
        if (string.IsNullOrEmpty(url))
            return new PositionLookupResult(PositionLookupOutcome.Error, null, "IamGroups:PositionLookupUrl is not configured.");

        try
        {
            var answer = await _api.SendJsonAsync(HttpMethod.Post, url, new { names = new[] { name } });

            if (!answer.Success)
            {
                var shown = answer.Body.Length > 300 ? answer.Body[..300] : answer.Body;
                return new PositionLookupResult(PositionLookupOutcome.Error, null, $"IAM answered {answer.StatusCode}: {shown}", answer.StatusCode);
            }
            return ParseBody(answer.Body, name, answer.StatusCode);
        }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException or JsonException)
        {
            _logger.LogWarning(ex, "Position lookup failed for {Name}", name);
            return new PositionLookupResult(PositionLookupOutcome.Error, null, ex.Message);
        }
    }

    internal static PositionLookupResult ParseBody(string body, string query, int status)
    {
        using var doc = JsonDocument.Parse(string.IsNullOrWhiteSpace(body) ? "null" : body);
        var el = doc.RootElement;

        // The answer is a list with one entry per name asked for; take the entry for this query.
        if (el.ValueKind == JsonValueKind.Array)
        {
            JsonElement? entry = null;
            foreach (var e in el.EnumerateArray())
            {
                if (e.ValueKind != JsonValueKind.Object) continue;
                entry ??= e;
                if (e.TryGetProperty("query", out var q) && q.ValueKind == JsonValueKind.String && Same(q.GetString(), query))
                {
                    entry = e;
                    break;
                }
            }
            if (entry == null) return new PositionLookupResult(PositionLookupOutcome.NotFound, null, "The answer has no entry for this name.", status);
            el = entry.Value;
        }

        // Anything else is not the documented shape: report it instead of treating it as "no such position".
        if (el.ValueKind != JsonValueKind.Object || !el.TryGetProperty("matches", out var matches) || matches.ValueKind != JsonValueKind.Array)
        {
            var text = body.Length > 300 ? body[..300] : body;
            return new PositionLookupResult(PositionLookupOutcome.Error, null, $"Unexpected answer: {text}", status);
        }

        var all = matches.EnumerateArray().Where(m => m.ValueKind == JsonValueKind.Object).ToList();
        if (all.Count == 0)
            return new PositionLookupResult(PositionLookupOutcome.NotFound, null, "No position with that name.", status);

        var chosen = all;
        if (all.Count > 1)
        {
            chosen = all.Where(m => m.TryGetProperty("name", out var n) && n.ValueKind == JsonValueKind.Object &&
                                    (Same(Str(n, "en"), query) || Same(Str(n, "am"), query))).ToList();
            if (chosen.Count != 1)
                return new PositionLookupResult(PositionLookupOutcome.Ambiguous, null, $"{all.Count} positions match that name.", status);
        }

        var id = Str(chosen[0], "id");
        return string.IsNullOrWhiteSpace(id)
            ? new PositionLookupResult(PositionLookupOutcome.Error, null, "The matching position has no id.", status)
            : new PositionLookupResult(PositionLookupOutcome.Found, id.Trim(), null, status);
    }

    private static string? Str(JsonElement obj, string name) =>
        obj.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() : null;

    private static bool Same(string? a, string? b) =>
        string.Equals(a?.Trim(), b?.Trim(), StringComparison.OrdinalIgnoreCase);
}

using System.IO;
using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using SPMigration.Models;

public class ApiService
{
    private readonly HttpClient _httpClient;
    private readonly IConfiguration _configuration;
    private readonly ILogger<ApiService> _logger;
    private readonly SemaphoreSlim _refreshLock = new(1, 1);
    private string? _token;
    private string? _refreshToken;

    public string? Token => _token;
    public string? RefreshToken => _refreshToken;

    public ApiService(HttpClient httpClient, IConfiguration configuration, ILogger<ApiService> logger)
    {
        _httpClient = httpClient;
        _configuration = configuration;
        _logger = logger;
    }

    public async Task<bool> LoginAsync()
    {
        try
        {
            var loginUrl = _configuration["Api:LoginUrl"];
            var loginData = new
            {
                email = _configuration["Api:User"],
                password = _configuration["Api:Password"]
            };

            var content = new StringContent(JsonSerializer.Serialize(loginData), Encoding.UTF8, "application/json");
            var response = await _httpClient.PostAsync(loginUrl, content);

            if (response.IsSuccessStatusCode)
            {
                var responseContent = await response.Content.ReadAsStringAsync();
                SaveTokens(responseContent);
                Console.WriteLine(_token);
                return true;
            }

            Console.WriteLine(_token);
            _logger.LogError("Login failed: {StatusCode}", response.StatusCode);
            return false;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error during login");
            return false;
        }
    }

    public async Task<bool> RefreshTokenAsync()
    {
        await _refreshLock.WaitAsync();
        try
        {
            var refreshUrl = _configuration["Api:RefreshTokenUrl"];
            if (string.IsNullOrEmpty(refreshUrl))
            {
                var loginUrl = _configuration["Api:LoginUrl"];
                if (!string.IsNullOrEmpty(loginUrl) && loginUrl.Contains("/auth/login"))
                {
                    refreshUrl = loginUrl.Replace("/auth/login", "/auth/refresh-token");
                }
            }
            else if (!Uri.IsWellFormedUriString(refreshUrl, UriKind.Absolute))
            {
                var loginUrl = _configuration["Api:LoginUrl"];
                if (Uri.TryCreate(loginUrl, UriKind.Absolute, out var baseUri))
                {
                    refreshUrl = new Uri(baseUri, refreshUrl).ToString();
                }
            }

            if (string.IsNullOrEmpty(refreshUrl))
            {
                _logger.LogError("RefreshTokenUrl is not configured in appsettings.json");
                return false;
            }

            if (string.IsNullOrEmpty(_refreshToken))
            {
                _logger.LogWarning("No refresh token stored. Attempting full login...");
                return await LoginAsync();
            }

            var payload = new
            {
                refreshToken = _refreshToken
            };

            var content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json");

            // Clear Authorization header for refresh request to avoid rejection if previous token expired
            _httpClient.DefaultRequestHeaders.Authorization = null;

            var response = await _httpClient.PostAsync(refreshUrl, content);

            if (response.IsSuccessStatusCode)
            {
                var responseContent = await response.Content.ReadAsStringAsync();
                SaveTokens(responseContent);
                _logger.LogInformation("✅ Token refreshed successfully.");
                return true;
            }

            var errorMsg = await response.Content.ReadAsStringAsync();
            _logger.LogError("Token refresh failed: {StatusCode} {Message}. Attempting full login fallback...", response.StatusCode, errorMsg);
            return await LoginAsync();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error during token refresh");
            return false;
        }
        finally
        {
            _refreshLock.Release();
        }
    }

    private void SaveTokens(string responseContent)
    {
        using var doc = JsonDocument.Parse(responseContent);
        var root = doc.RootElement;
        var target = root;

        if (!root.TryGetProperty("token", out _) &&
            root.TryGetProperty("data", out var dataProp) &&
            dataProp.ValueKind == JsonValueKind.Object)
        {
            target = dataProp;
        }

        if (target.TryGetProperty("token", out var tokenProp))
        {
            _token = tokenProp.GetString();
        }
        else if (target.TryGetProperty("Token", out var tokenPascal))
        {
            _token = tokenPascal.GetString();
        }

        if (target.TryGetProperty("refreshToken", out var refreshProp))
        {
            _refreshToken = refreshProp.GetString();
        }
        else if (target.TryGetProperty("refresh_token", out var refreshSnake))
        {
            _refreshToken = refreshSnake.GetString();
        }
        else if (target.TryGetProperty("RefreshToken", out var refreshPascal))
        {
            _refreshToken = refreshPascal.GetString();
        }

        if (!string.IsNullOrEmpty(_token))
        {
            _httpClient.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", _token);
        }
    }

    private bool IsJwtTokenError(HttpStatusCode statusCode, string? responseContent)
    {
        if (statusCode == HttpStatusCode.Unauthorized)
        {
            return true;
        }

        if (!string.IsNullOrWhiteSpace(responseContent))
        {
            var content = responseContent.ToLowerInvariant();
            if (content.Contains("jwt token error") ||
                content.Contains("jwt expired") ||
                content.Contains("token expired") ||
                content.Contains("invalid token") ||
                content.Contains("token is invalid") ||
                content.Contains("jwt error") ||
                (content.Contains("jwt") && content.Contains("error")))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// POSTs a JSON body to an arbitrary IAM URL with the bearer token, logging in / refreshing when needed.
    /// </summary>
    public async Task<(bool Success, int StatusCode, string Body)> PostJsonAsync(string url, object payload)
    {
        var json = JsonSerializer.Serialize(payload, new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
        });

        if (string.IsNullOrEmpty(_token) && !await LoginAsync())
            return (false, 401, "Login to IAM failed.");

        async Task<HttpResponseMessage> SendAsync()
        {
            using var request = new HttpRequestMessage(HttpMethod.Post, url)
            {
                Content = new StringContent(json, Encoding.UTF8, "application/json")
            };
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _token);
            return await _httpClient.SendAsync(request);
        }

        var response = await SendAsync();
        var body = await response.Content.ReadAsStringAsync();

        if (!response.IsSuccessStatusCode && IsJwtTokenError(response.StatusCode, body) && await RefreshTokenAsync())
        {
            response = await SendAsync();
            body = await response.Content.ReadAsStringAsync();
        }

        return (response.IsSuccessStatusCode, (int)response.StatusCode, body);
    }

    public async Task<bool> MigrateRecordAsync(MigrationPayload payload)
    {
        try
        {
            var migrateUrl = _configuration["Api:MigrateUrl"];

            var json = JsonSerializer.Serialize(payload, new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                WriteIndented = false
            });
            var options = new JsonSerializerOptions { WriteIndented = true };
            var prettyJson = JsonSerializer.Serialize(JsonSerializer.Deserialize<object>(json), options);
            Console.WriteLine(JsonSerializer.Serialize(payload, new JsonSerializerOptions
            {
                WriteIndented = true
            }));

            var content = new StringContent(json, Encoding.UTF8, "application/json");
            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _token);

            var response = await _httpClient.PostAsync(migrateUrl, content);
            var responseContent = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode || IsJwtTokenError(response.StatusCode, responseContent))
            {
                if (IsJwtTokenError(response.StatusCode, responseContent))
                {
                    _logger.LogWarning("⚠️ JWT token error encountered in MigrateRecordAsync ({StatusCode}: {Message}). Refreshing token...", response.StatusCode, responseContent);
                    if (await RefreshTokenAsync())
                    {
                        var retryContent = new StringContent(json, Encoding.UTF8, "application/json");
                        _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _token);
                        response = await _httpClient.PostAsync(migrateUrl, retryContent);
                        responseContent = await response.Content.ReadAsStringAsync();
                    }
                }

                if (!response.IsSuccessStatusCode)
                {
                    _logger.LogError("❌ Failed to migrate record: {StatusCode} {Message}", response.StatusCode, responseContent);
                }
            }

            return response.IsSuccessStatusCode;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error migrating record");
            return false;
        }
    }

    public async Task<bool> MigrateExternalRecordAsync(IncomingRecordMigrationPayload payload)
    {
        try
        {
            var migrateUrl = _configuration["Api:MigrateIncomingUrl"];

            var json = JsonSerializer.Serialize(payload, new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                WriteIndented = false
            });
            var options = new JsonSerializerOptions { WriteIndented = true };
            var prettyJson = JsonSerializer.Serialize(JsonSerializer.Deserialize<object>(json), options);

            Console.WriteLine("📦 Payload:");
            Console.WriteLine(prettyJson);

            var content = new StringContent(json, Encoding.UTF8, "application/json");
            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _token);

            var response = await _httpClient.PostAsync(migrateUrl, content);
            var responseContent = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode || IsJwtTokenError(response.StatusCode, responseContent))
            {
                if (IsJwtTokenError(response.StatusCode, responseContent))
                {
                    _logger.LogWarning("⚠️ JWT token error encountered in MigrateExternalRecordAsync ({StatusCode}: {Message}). Refreshing token...", response.StatusCode, responseContent);
                    if (await RefreshTokenAsync())
                    {
                        var retryContent = new StringContent(json, Encoding.UTF8, "application/json");
                        _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _token);
                        response = await _httpClient.PostAsync(migrateUrl, retryContent);
                        responseContent = await response.Content.ReadAsStringAsync();
                    }
                }

                if (!response.IsSuccessStatusCode)
                {
                    _logger.LogError("❌ Failed to migrate external record: {StatusCode} {Message}", response.StatusCode, responseContent);
                }
            }

            return response.IsSuccessStatusCode;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error migrating record");
            return false;
        }
    }

    public async Task<bool> UpdateRecordFilesAsync(FileInfoMigrationDto fileInfoDto)
    {
        try
        {
            var updateUrl = $"{_configuration["Api:UpdateFilesUrl"]}{fileInfoDto.Id}";
            var payload = new
            {
                fileInfo = fileInfoDto.FileInfo,
                attachments = fileInfoDto.Attachments
            };

            var json = JsonSerializer.Serialize(payload, new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                WriteIndented = false
            });

            var options = new JsonSerializerOptions { WriteIndented = true };
            var prettyJson = JsonSerializer.Serialize(JsonSerializer.Deserialize<object>(json), options);

            Console.WriteLine("📎 File Update Payload:");
            Console.WriteLine(prettyJson);

            var content = new StringContent(json, Encoding.UTF8, "application/json");
            _httpClient.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", _token);

            var response = await _httpClient.PatchAsync(updateUrl, content);
            var responseContent = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode || IsJwtTokenError(response.StatusCode, responseContent))
            {
                if (IsJwtTokenError(response.StatusCode, responseContent))
                {
                    _logger.LogWarning("⚠️ JWT token error encountered in UpdateRecordFilesAsync ({StatusCode}: {Message}). Refreshing token...", response.StatusCode, responseContent);
                    if (await RefreshTokenAsync())
                    {
                        var retryContent = new StringContent(json, Encoding.UTF8, "application/json");
                        _httpClient.DefaultRequestHeaders.Authorization =
                            new AuthenticationHeaderValue("Bearer", _token);
                        response = await _httpClient.PatchAsync(updateUrl, retryContent);
                        responseContent = await response.Content.ReadAsStringAsync();
                    }
                }
            }

            if (response.IsSuccessStatusCode)
            {
                _logger.LogInformation("✅ File details updated for record {Id}", fileInfoDto.Id);
                return true;
            }

            _logger.LogError("❌ Failed to update file details for record {Id}: {StatusCode} {Message}",
                fileInfoDto.Id, response.StatusCode, responseContent);

            return false;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error updating file details for record {Id}", fileInfoDto.Id);
            return false;
        }
    }
}

using System.Text.Json;
using Microsoft.AspNetCore.Mvc;
using SPMigration.Models;
using System.Text.Json;
namespace SPMigration.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class MigrationController : ControllerBase
    {
        private readonly MigrationService _migrationService;
        private readonly ILogger<MigrationController> _logger;

        public MigrationController(MigrationService migrationService, ApiService apiService, ILogger<MigrationController> logger)
        {
            _migrationService = migrationService;

            _logger = logger;
        }

        /// <summary>
        /// Run the migration process for a specific range of IDs.
        /// Optional recordType (case-insensitive, comma-separate several) migrates only records of that Record_Type and skips the rest.
        /// Example: POST /api/migration/run?startId=0&endId=300&recordType=Internal
        /// </summary>
        [HttpPost("run")]
        public async Task<IActionResult> RunMigration(
        [FromQuery] int startId,
        [FromQuery] int endId,
        IFormFile? file = null,
        [FromQuery] string? recordType = null)
        {
            if (startId < 0 || endId <= startId)
            {
                return BadRequest(new
                {
                    error = "Invalid range. Ensure startId >= 0 and endId > startId."
                });
            }

            try
            {
                _logger.LogInformation(
                    "▶ Triggered migration via API for range {StartId}-{EndId}.",
                    startId,
                    endId
                );

                var (userMappings, mappingError) = await ReadUserMappingsAsync(file);
                if (mappingError != null)
                    return mappingError;

                await _migrationService.RunMigrationAsync(
                    startId,
                    endId,
                    userMappings,
                    recordType
                );

                return Ok(new
                {
                    message = $"Migration request for range {startId}-{endId} finished."
                });
            }
            catch (JsonException ex)
            {
                return BadRequest(new
                {
                    error = $"Invalid JSON format: {ex.Message}"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(
                    ex,
                    "❌ Migration failed for range {StartId}-{EndId}.",
                    startId,
                    endId
                );

                return StatusCode(500, new
                {
                    error = ex.Message
                });
            }
        }

        /// <summary>
        /// Run the migration process for a specific range of IDs.
        /// Example: GET /api/migration/run?startId=0&endId=300
        /// </summary>
        [HttpGet("FindAndStoreAllUsersInJson")]
        public async Task<IActionResult> FindAndStoreAllUsersInJson([FromQuery] int startId, [FromQuery] int endId)
        {
            if (startId < 0 || endId <= startId)
            {
                return BadRequest(new { error = "Invalid range. Ensure startId >= 0 and endId > startId." });
            }

            try
            {
                _logger.LogInformation("▶ Triggered migration via API for range {StartId}-{EndId}.", startId, endId);

                await _migrationService.FindAndStoreAllUsersInJson(startId, endId);

                return Ok(new { message = $"Migration for range {startId}-{endId} completed successfully." });
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                _logger.LogError(ex, "❌ Migration failed for range {StartId}-{EndId}.", startId, endId);
                return StatusCode(500, new { error = ex.Message });
            }
        }

        /// <summary>
        /// Run the migration process for a specific range of IDs.
        /// Example: GET /api/migration/run?startId=0&endId=300
        /// </summary>
        [HttpPost("runForIds")]
        public async Task<IActionResult> RunMigrationForIds([FromBody] List<string> recordIds)
        {
            try
            {
                _logger.LogInformation("▶ Triggered migration via API for range {StartId}-{EndId}.");

                await _migrationService.RunMigrationForRecordIdsAsync(recordIds);

                return Ok(new { message = $"Migration for range {recordIds[0]}_{recordIds[recordIds.Count() - 1]} completed successfully." });
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                _logger.LogError(ex, "❌ Migration failed for range {StartId}-{EndId}.");
                return StatusCode(500, new { error = ex.Message });
            }
        }
        /// <summary>Migrates the given Records/Incoming folder IDs with an optional user-mapping JSON file (multipart/form-data). ids is JSON text: { "ids": [9141, 9174] }.</summary>
        [HttpPost("incoming-by-ids")]
        public async Task<IActionResult> MigrateIncomingByIds([FromForm] string ids, IFormFile? file = null)
        {
            var (idList, idsError) = ParseIds(ids);
            if (idsError != null)
                return idsError;

            var (userMappings, mappingError) = await ReadUserMappingsAsync(file);
            if (mappingError != null)
                return mappingError;

            try
            {
                await _migrationService.RunIncomingForFolderIdsAsync(idList!, userMappings);
                return Ok(new { message = $"Incoming migration finished for {idList!.Count} id(s)." });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "❌ Incoming migration by ids failed");
                return StatusCode(500, new { error = ex.Message });
            }
        }

        /// <summary>Migrates the given User_Records IDs with an optional user-mapping JSON file (multipart/form-data). ids is JSON text: { "ids": [9141, 9174] }.</summary>
        [HttpPost("user-records-by-ids")]
        public async Task<IActionResult> MigrateUserRecordsByIds([FromForm] string ids, IFormFile? file = null)
        {
            var (idList, idsError) = ParseIds(ids);
            if (idsError != null)
                return idsError;

            var (userMappings, mappingError) = await ReadUserMappingsAsync(file);
            if (mappingError != null)
                return mappingError;

            try
            {
                await _migrationService.RunMigrationForRecordIdsAsync(
                    idList!.Select(i => i.ToString()).ToList(), userMappings);
                return Ok(new { message = $"User_Records migration finished for {idList!.Count} id(s)." });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "❌ User_Records migration by ids failed");
                return StatusCode(500, new { error = ex.Message });
            }
        }

        /// <summary>
        /// Run the document-library migration for a specific inclusive range of folder IDs.
        /// Optionally accepts a JSON user-mapping file as multipart/form-data.
        /// </summary>
        [HttpPost("runForIdsDocLibrary")]
        public async Task<IActionResult> RunMigrationForIdsDocLibrary(
            [FromQuery] int startId,
            [FromQuery] int endId,
            IFormFile? file = null)
        {
            if (startId < 0 || endId <= startId)
            {
                return BadRequest(new
                {
                    error = "Invalid range. Ensure startId >= 0 and endId > startId."
                });
            }

            try
            {
                _logger.LogInformation(
                    "▶ Triggered document-library migration via API for range {StartId}-{EndId}.",
                    startId,
                    endId);

                var (userMappings, mappingError) = await ReadUserMappingsAsync(file);
                if (mappingError != null)
                    return mappingError;

                await _migrationService.RunMigrationForRecordIdsDocLibraryAsync(
                    startId,
                    endId,
                    userMappings);

                return Ok(new
                {
                    message = $"Document-library migration for range {startId}-{endId} finished."
                });
            }
            catch (JsonException ex)
            {
                return BadRequest(new { error = $"Invalid JSON format: {ex.Message}" });
            }
            catch (Exception ex)
            {
                _logger.LogError(
                    ex,
                    "❌ Document-library migration failed for range {StartId}-{EndId}.",
                    startId,
                    endId);
                return StatusCode(500, new { error = ex.Message });
            }
        }
        /// <summary>
        /// </summary>
        [HttpGet("runForExternalIncoming")]
        public async Task<IActionResult> RunDocumentLibraryMigrationAsync()
        {

            try
            {

                await _migrationService.RunDocumentLibraryMigrationAsync();

                return Ok(new { message = $"Migration completed successfully." });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "❌ Migration failed");
                return StatusCode(500, new { error = ex.Message });
            }
        }

        
        /// <summary>
        /// </summary>
        [HttpGet("FindAndStoreAllUsersInJsonForDocumentLibary")]
        public async Task<IActionResult> FindAndStoreAllUsersInJsonForDocumentLibary()
        {
            try
            {

                await _migrationService.FindAndStoreAllUsersInJsonForDocumentLibary();

                return Ok(new { message = $"Migration completed successfully." });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "❌ Migration failed");
                return StatusCode(500, new { error = ex.Message });
            }
        }

        /// <summary>
        /// User_Records ready for migration. Already-migrated items are counted but omitted unless includeMigrated=true.
        /// limit caps the ids returned (default 100; 0 = no cap); totalReady always has the full number.
        /// Swagger UI freezes on very large bodies, so for the full list use download=true (saves it as a JSON file).
        /// recordType filters on the SharePoint Record_Type (case-insensitive, comma-separate several); availableRecordTypes in the response lists the valid values.
        /// Example: GET /api/migration/ready/user-records?limit=100&amp;recordType=Internal
        /// </summary>
        [HttpGet("ready/user-records")]
        public async Task<IActionResult> GetReadyUserRecords([FromQuery] bool includeMigrated = false, [FromQuery] int limit = 100, [FromQuery] bool download = false, [FromQuery] string? recordType = null)
        {
            try
            {
                return ReadyResponse(await _migrationService.GetReadyUserRecordsAsync(includeMigrated, recordType), limit, download, "ready-user-records.json");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "❌ Failed to list User_Records ready for migration");
                return StatusCode(500, new { error = ex.Message });
            }
        }

        /// <summary>
        /// Records/Incoming folders ready for migration. Same parameters as ready/user-records.
        /// Example: GET /api/migration/ready/incoming?limit=100
        /// </summary>
        [HttpGet("ready/incoming")]
        public async Task<IActionResult> GetReadyIncomingRecords([FromQuery] bool includeMigrated = false, [FromQuery] int limit = 100, [FromQuery] bool download = false)
        {
            try
            {
                return ReadyResponse(await _migrationService.GetReadyIncomingRecordsAsync(includeMigrated), limit, download, "ready-incoming-records.json");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "❌ Failed to list Records/Incoming ready for migration");
                return StatusCode(500, new { error = ex.Message });
            }
        }

        // Parses the "ids" form field: JSON such as {"ids":[9141,9174]} (a bare array [9141,9174] also works).
        private static (List<int>? Ids, IActionResult? Error) ParseIds(string? ids)
        {
            const string msg = "ids must be JSON like { \"ids\": [9141, 9174] } with at least one number.";
            if (string.IsNullOrWhiteSpace(ids))
                return (null, new BadRequestObjectResult(new { error = msg }));
            try
            {
                using var doc = JsonDocument.Parse(ids, new JsonDocumentOptions { AllowTrailingCommas = true });
                var root = doc.RootElement;
                if (root.ValueKind == JsonValueKind.Object)
                {
                    var prop = root.EnumerateObject()
                        .FirstOrDefault(p => p.Name.Equals("ids", StringComparison.OrdinalIgnoreCase));
                    if (prop.Value.ValueKind != JsonValueKind.Array)
                        return (null, new BadRequestObjectResult(new { error = msg }));
                    root = prop.Value;
                }
                var list = JsonSerializer.Deserialize<List<int>>(root.GetRawText());
                if (list == null || list.Count == 0)
                    return (null, new BadRequestObjectResult(new { error = msg }));
                return (list, null);
            }
            catch (JsonException ex)
            {
                return (null, new BadRequestObjectResult(new { error = $"{msg} {ex.Message}" }));
            }
        }

        // Reads the optional user-mapping JSON upload. No file → (null, null), i.e. run without mappings.
        private async Task<(List<UserMapping>? Mappings, IActionResult? Error)> ReadUserMappingsAsync(IFormFile? file)
        {
            if (file == null || file.Length == 0)
            {
                _logger.LogInformation("No user mapping file provided. Proceeding without mappings.");
                return (null, null);
            }

            if (!file.FileName.EndsWith(".json", StringComparison.OrdinalIgnoreCase))
                return (null, BadRequest(new { error = "Only JSON files are supported." }));

            using var reader = new StreamReader(file.OpenReadStream());
            var fileContent = await reader.ReadToEndAsync();

            if (string.IsNullOrWhiteSpace(fileContent))
                return (null, BadRequest(new { error = "JSON file is empty." }));

            List<UserMapping>? userMappings;
            try
            {
                userMappings = JsonSerializer.Deserialize<List<UserMapping>>(
                    fileContent,
                    new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
            }
            catch (JsonException ex)
            {
                return (null, BadRequest(new { error = $"Invalid JSON format: {ex.Message}" }));
            }

            if (userMappings == null)
                return (null, BadRequest(new { error = "JSON must contain a valid user mapping array." }));

            _logger.LogInformation("📁 Loaded {Count} user mappings.", userMappings.Count);
            return (userMappings, null);
        }

        // download=true returns the whole list (limit ignored) as a JSON file so Swagger doesn't have to render it.
        private IActionResult ReadyResponse(ReadyRecordsResult result, int limit, bool download, string fileName)
        {
            result.TotalReady = result.Ids.Count;
            if (download)
            {
                var bytes = JsonSerializer.SerializeToUtf8Bytes(result, new JsonSerializerOptions
                {
                    WriteIndented = true,
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                });
                return File(bytes, "application/json", fileName);
            }

            if (limit > 0)
            {
                result.Ids = result.Ids.Take(limit).ToList();
                result.FailedIds = result.FailedIds.Take(limit).ToList();
                result.FailedDetails = result.FailedDetails.Take(limit).ToList();
                result.FailedByCategory = result.FailedByCategory.ToDictionary(k => k.Key, k => k.Value.Take(limit).ToList());
                result.NotMigratedIds = result.NotMigratedIds.Take(limit).ToList();
            }
            return Ok(result);
        }

        /// <summary>
        /// Retries the records in an errors file. Optionally accepts a JSON user-mapping file as multipart/form-data.
        /// </summary>
        [HttpPost("retry-from-errors")]
        public async Task<IActionResult> RetryFromErrors([FromQuery] string fileName, IFormFile? file = null)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                return BadRequest(new { error = "File name must be provided." });

            var (userMappings, mappingError) = await ReadUserMappingsAsync(file);
            if (mappingError != null)
                return mappingError;

            try
            {

                await _migrationService.RetryErrors(fileName, userMappings);

                return Ok(new { message = $"Migration completed successfully." });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "❌ Migration failed");
                return StatusCode(500, new { error = ex.Message });
            }
        }

        /// <summary>
        /// Retries the folders in a doc-library errors file. Optionally accepts a JSON user-mapping file as multipart/form-data.
        /// </summary>
        [HttpPost("retry-from-errors-doc-lib")]
        public async Task<IActionResult> RetryFromErrorsDocLib([FromQuery] string fileName, IFormFile? file = null)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                return BadRequest(new { error = "File name must be provided." });

            var (userMappings, mappingError) = await ReadUserMappingsAsync(file);
            if (mappingError != null)
                return mappingError;

            try
            {

                await _migrationService.RetryErrorsDocLib(fileName, userMappings);

                return Ok(new { message = $"Migration completed successfully." });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "❌ Migration failed");
                return StatusCode(500, new { error = ex.Message });
            }
        }

        /// <summary>
        /// Groups the incoming items by Path and attaches AD members to each group.
        /// </summary>
        [HttpPost("groups-with-members")]
        public async Task<IActionResult> GetGroupsWithMembers(
            [FromBody] List<GroupItem> groups,
            [FromServices] AdExplorerService adService,
            CancellationToken ct)
        {
            if (groups == null || groups.Count == 0)
                return BadRequest(new { error = "Request body must be a non-empty list of groups." });

            try
            {
                return Ok(await adService.GetMembersByPathAsync(groups, ct));
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "❌ Failed to get groups with members");
                return StatusCode(500, new { error = ex.Message });
            }
        }

        /// <summary>
        /// Creates the groups (with members and subgroups) from the attached JSON file on the IAM side.
        /// unitId is applied to every group. Use dryRun=true to download the generated payloads as iam-groups-payload.json instead of posting.
        /// </summary>
        [HttpPost("iam-groups")]
        public async Task<IActionResult> CreateIamGroups(
            [FromServices] IamGroupMigrationService iamGroups,
            IFormFile file,
            [FromQuery] string? unitId = null,
            [FromQuery] bool dryRun = true,
            CancellationToken ct = default)
        {
            if (file == null || file.Length == 0)
                return BadRequest(new { error = "Attach the groups JSON file." });
            if (!dryRun && string.IsNullOrWhiteSpace(unitId))
                return BadRequest(new { error = "unitId is required when dryRun=false." });

            try
            {
                using var reader = new StreamReader(file.OpenReadStream());
                var json = await reader.ReadToEndAsync(ct);

                var input = JsonSerializer.Deserialize<List<UpdatedPathEntry>>(json,
                    new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                if (input == null || input.Count == 0)
                    return BadRequest(new { error = "Input contains no groups." });

                var report = await iamGroups.RunAsync(input, unitId, dryRun, ct);

                if (dryRun)
                {
                    // Dry run: return the generated payloads as a downloadable JSON file.
                    var bytes = JsonSerializer.SerializeToUtf8Bytes(report.Payloads, new JsonSerializerOptions
                    {
                        WriteIndented = true,
                        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                        DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull,
                        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                    });
                    return File(bytes, "application/json", "iam-groups-payload.json");
                }

                return Ok(report);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "❌ IAM group creation failed");
                return StatusCode(500, new { error = ex.Message });
            }
        }
    }

}

public class RetryRequest
{
    public string RecordsList { get; set; } = string.Empty;
    public string TasksList { get; set; } = string.Empty;
    public List<int> RecordIds { get; set; } = new();
}

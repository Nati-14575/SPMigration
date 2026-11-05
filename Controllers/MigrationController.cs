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
        /// Example: GET /api/migration/run?startId=0&endId=300
        /// </summary>
        [HttpPost("run")]
        public async Task<IActionResult> RunMigration(
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
                    "▶ Triggered migration via API for range {StartId}-{EndId}.",
                    startId,
                    endId
                );

                List<UserMapping>? userMappings = null;

                // Only deserialize when a file is provided
                if (file != null && file.Length > 0)
                {
                    if (!file.FileName.EndsWith(
                        ".json",
                        StringComparison.OrdinalIgnoreCase))
                    {
                        return BadRequest(new
                        {
                            error = "Only JSON files are supported."
                        });
                    }

                    using var reader = new StreamReader(file.OpenReadStream());

                    var fileContent = await reader.ReadToEndAsync();

                    if (string.IsNullOrWhiteSpace(fileContent))
                    {
                        return BadRequest(new
                        {
                            error = "JSON file is empty."
                        });
                    }

                    userMappings = JsonSerializer.Deserialize<List<UserMapping>>(
                        fileContent,
                        new JsonSerializerOptions
                        {
                            PropertyNameCaseInsensitive = true
                        }
                    );

                    if (userMappings == null)
                    {
                        return BadRequest(new
                        {
                            error = "JSON must contain a valid user mapping array."
                        });
                    }

                    _logger.LogInformation(
                        "📁 Loaded {Count} user mappings.",
                        userMappings.Count
                    );
                }
                else
                {
                    _logger.LogInformation(
                        "No user mapping file provided. Proceeding without mappings."
                    );
                }

                await _migrationService.RunMigrationAsync(
                    startId,
                    endId,
                    userMappings
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

                List<UserMapping>? userMappings = null;

                if (file != null && file.Length > 0)
                {
                    if (!file.FileName.EndsWith(".json", StringComparison.OrdinalIgnoreCase))
                    {
                        return BadRequest(new { error = "Only JSON files are supported." });
                    }

                    using var reader = new StreamReader(file.OpenReadStream());
                    var fileContent = await reader.ReadToEndAsync();

                    if (string.IsNullOrWhiteSpace(fileContent))
                    {
                        return BadRequest(new { error = "JSON file is empty." });
                    }

                    userMappings = JsonSerializer.Deserialize<List<UserMapping>>(
                        fileContent,
                        new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

                    if (userMappings == null)
                    {
                        return BadRequest(new
                        {
                            error = "JSON must contain a valid user mapping array."
                        });
                    }

                    _logger.LogInformation("📁 Loaded {Count} user mappings.", userMappings.Count);
                }

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

        [HttpPost("retry-from-errors")]
        public async Task<IActionResult> RetryFromErrors([FromQuery] string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                return BadRequest(new { error = "File name must be provided." });

            try
            {

                await _migrationService.RetryErrors(fileName);

                return Ok(new { message = $"Migration completed successfully." });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "❌ Migration failed");
                return StatusCode(500, new { error = ex.Message });
            }
        }

        [HttpPost("retry-from-errors-doc-lib")]
        public async Task<IActionResult> RetryFromErrorsDocLib([FromQuery] string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                return BadRequest(new { error = "File name must be provided." });

            try
            {

                await _migrationService.RetryErrorsDocLib(fileName);

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

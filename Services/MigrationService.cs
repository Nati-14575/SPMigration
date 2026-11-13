using System.Net.Http.Json;
using System.Text.Json;
using Microsoft.AspNetCore.Http.HttpResults;
using Microsoft.Extensions.Logging;
using SPMigration.Models;

public class MigrationService
{
    private readonly SharePointService _sharePointService;
    private readonly MinIOService _minioService;
    private readonly ApiService _apiService;
    private readonly IConfiguration _configuration;
    private readonly ILogger<MigrationService> _logger;

    public MigrationService(
        SharePointService sharePointService,
        MinIOService minioService,
        ApiService apiService,
        IConfiguration configuration,
        ILogger<MigrationService> logger)
    {
        _sharePointService = sharePointService;
        _minioService = minioService;
        _apiService = apiService;
        _configuration = configuration;
        _logger = logger;
    }

    /// <param name="recordType">Optional Record_Type filter (case-insensitive); comma-separate several. Records of other types are skipped.</param>
    public async Task RunMigrationAsync(int? startId = null, int? endId = null, List<UserMapping> userMappings = null, string? recordType = null)
    {
        var startTime = DateTime.UtcNow;
        var processedCount = 0;
        var errorCount = 0;

        _logger.LogInformation("🚀 Starting migration at {StartTime}", startTime);

        try
        {
            if (!await _apiService.LoginAsync())
            {
                _logger.LogError("❌ Failed to login to API, aborting migration");
                return;
            }

            var recordsList = _configuration["SharePoint:RecordsListName"] ?? "User_Records";
            var tasksList = _configuration["SharePoint:TasksListName"] ?? "User_Records_Associated_Task";
            var internalTasksList = _configuration["SharePoint:InternalTasksListName"] ?? "Internal_Records_Associated_Task";
            var internalList = _configuration["SharePoint:InternalListName"] ?? "Internal_Records";
            var rowLimit = _configuration.GetValue<int?>("SharePoint:RowLimit") ?? 500;
            var lastProcessedId = await GetLastProcessedIdAsync();
            var wantedTypes = ParseRecordTypes(recordType);
            if (wantedTypes.Count > 0)
                _logger.LogInformation("🔎 Only migrating Record_Type: {Types}", string.Join(",", wantedTypes));

            List<int> recordIds;
            string runMode;

            if (startId.HasValue && endId.HasValue)
            {
                runMode = "Range";
                _logger.LogInformation("▶ Running migration for range {Start}–{End}", startId, endId);
                recordIds = Enumerable.Range(startId.Value, endId.Value - startId.Value + 1).ToList();
            }
            else
            {
                runMode = "Resume";
                _logger.LogInformation("▶ Resuming from ID > {LastId}", lastProcessedId);
                recordIds = await _sharePointService.GetRecordIdsAsync(recordsList, lastProcessedId, rowLimit);
                startId = lastProcessedId + 1;
                endId = lastProcessedId + recordIds.Count;
            }

            _logger.LogInformation("📥 Loaded {Count} records to process", recordIds.Count);

            foreach (var id in recordIds)
            {
                try
                {
                    var record = await _sharePointService.GetRecordByIdAsync(recordsList, id, userMappings);

                    if (record == null)
                    {
                        _logger.LogWarning("⚠️ Record {Id} not found in SharePoint.", id);
                        continue;
                    }

                    if (wantedTypes.Count > 0 && !wantedTypes.Contains(RecordTypeKey(record.Record_Type)))
                    {
                        _logger.LogInformation("⏭ Skipping record {Id}: Record_Type {Type} is not in the requested filter.", id, record.Record_Type);
                        continue;
                    }

                    if (record.Record_Type != "InternalMemo")
                    {
                        var recordOfficerRecord = await _sharePointService.GetRecordOfficerDetailsByIdAsync(id, userMappings);
                        if(recordOfficerRecord == null)
                        {
                            continue;
                        }
                        record.Reference_Number = recordOfficerRecord.Reference_Number;
                        record.DateOfDispatch = recordOfficerRecord.DateOfDispatch;
                        record.DispatchedBy = recordOfficerRecord.DispatchedBy;
                        record.RecordId = recordOfficerRecord.ID;
                    }

                    await ProcessRecordAsync(record, recordsList, tasksList, internalList, internalTasksList, userMappings);
                    await SaveCheckpointAsync(record.ID);
                    processedCount++;

                    var success = new MigrationSuccess
                    {
                        RecordId = record.ID,
                        RecordsList = recordsList,
                        TasksList = tasksList,
                        Timestamp = DateTime.UtcNow
                    };

                    await LogSuccessToFileAsync(success,
                        $"{startId}_{endId}_success.json");
                }
                catch (Exception ex)
                {
                    errorCount++;
                    _logger.LogError(ex, "❌ Error processing record {RecordId}", id);

                    var payloadJson = (ex is MigrationException mex && mex.Payload != null)
                        ? JsonSerializer.Serialize(mex.Payload, new JsonSerializerOptions { WriteIndented = true })
                        : null;

                    var error = new MigrationError
                    {
                        RecordId = id,
                        RecordsList = recordsList,
                        TasksList = tasksList,
                        InternalList = internalList,
                        InternalTasksList = internalTasksList,
                        ErrorMessage = ex.ToString(),
                        ErrorType = ErrorTypeOf(ex),
                        Payload = payloadJson,
                        Timestamp = DateTime.UtcNow
                    };

                    await LogErrorToFileAsync(error, $"{startId}_{endId}_errors.json");
                }
            }

            var endTime = DateTime.UtcNow;
            _logger.LogInformation("✅ Migration completed successfully at {EndTime}. Duration: {Duration} minutes",
                endTime, (endTime - startTime).TotalMinutes);

            // 📌 Save run log with range info
            var runLog = new MigrationRunLog
            {
                StartTime = startTime,
                EndTime = endTime,
                TotalProcessed = processedCount,
                TotalErrors = errorCount,
                StartId = startId,
                EndId = endId,
                Mode = runMode
            };

            await LogRunAsync(runLog, "migration_runs.json");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "💥 Migration failed unexpectedly.");
        }
    } 
    public async Task FindAndStoreAllUsersInJson(int? startId = null, int? endId = null, List<UserMapping> userMappings = null)
    {
        var startTime = DateTime.UtcNow;
        var processedCount = 0;
        var errorCount = 0;

        _logger.LogInformation("🚀 Starting migration at {StartTime}", startTime);

        try
        {
            if (!await _apiService.LoginAsync())
            {
                _logger.LogError("❌ Failed to login to API, aborting migration");
                return;
            }

            var recordsList = _configuration["SharePoint:RecordsListName"] ?? "User_Records";
            var tasksList = _configuration["SharePoint:TasksListName"] ?? "User_Records_Associated_Task";
            var internalTasksList = _configuration["SharePoint:InternalTasksListName"] ?? "Internal_Records_Associated_Task";
            var internalList = _configuration["SharePoint:InternalListName"] ?? "Internal_Records";
            var rowLimit = _configuration.GetValue<int?>("SharePoint:RowLimit") ?? 500;
            var lastProcessedId = await GetLastProcessedIdAsync();

            List<int> recordIds;
            string runMode;

            if (startId.HasValue && endId.HasValue)
            {
                runMode = "Range";
                _logger.LogInformation("▶ Running migration for range {Start}–{End}", startId, endId);
                recordIds = Enumerable.Range(startId.Value, endId.Value - startId.Value + 1).ToList();
            }
            else
            {
                runMode = "Resume";
                _logger.LogInformation("▶ Resuming from ID > {LastId}", lastProcessedId);
                recordIds = await _sharePointService.GetRecordIdsAsync(recordsList, lastProcessedId, rowLimit);
                startId = lastProcessedId + 1;
                endId = lastProcessedId + recordIds.Count;
            }

            _logger.LogInformation("📥 Loaded {Count} records to process", recordIds.Count);

            foreach (var id in recordIds)
            {
                try
                {
                    var record = await _sharePointService.GetRecordByIdAsync(recordsList, id);

                    if (record == null)
                    {
                        _logger.LogWarning("⚠️ Record {Id} not found in SharePoint.", id);
                        continue;
                    }

                    if (record.Record_Type != "InternalMemo")
                    {
                        var recordOfficerRecord = await _sharePointService.GetRecordOfficerDetailsByIdAsync(id);
                        if(recordOfficerRecord == null)
                        {
                            continue;
                        }
                        record.Reference_Number = recordOfficerRecord.Reference_Number;
                        record.DateOfDispatch = recordOfficerRecord.DateOfDispatch;
                        record.DispatchedBy = recordOfficerRecord.DispatchedBy;
                        record.RecordId = recordOfficerRecord.ID;
                    }

                    await ProcessRecordAsync(record, recordsList, tasksList, internalList, internalTasksList, userMappings,  false);
                    
                }
                catch (Exception ex)
                {
                    errorCount++;
                    _logger.LogError(ex, "❌ Error processing record {RecordId}", id);

                    var payloadJson = (ex is MigrationException mex && mex.Payload != null)
                        ? JsonSerializer.Serialize(mex.Payload, new JsonSerializerOptions { WriteIndented = true })
                        : null;

                    var error = new MigrationError
                    {
                        RecordId = id,
                        RecordsList = recordsList,
                        TasksList = tasksList,
                        InternalList = internalList,
                        InternalTasksList = internalTasksList,
                        ErrorMessage = ex.ToString(),
                        ErrorType = ErrorTypeOf(ex),
                        Payload = payloadJson,
                        Timestamp = DateTime.UtcNow
                    };

                    await LogErrorToFileAsync(error, $"{startId}_{endId}_errors.json");
                }
            }

            var endTime = DateTime.UtcNow;
            _logger.LogInformation("✅ Migration completed successfully at {EndTime}. Duration: {Duration} minutes",
                endTime, (endTime - startTime).TotalMinutes);

            // 📌 Save run log with range info
            var runLog = new MigrationRunLog
            {
                StartTime = startTime,
                EndTime = endTime,
                TotalProcessed = processedCount,
                TotalErrors = errorCount,
                StartId = startId,
                EndId = endId,
                Mode = runMode
            };

            await LogRunAsync(runLog, "migration_runs.json");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "💥 Migration failed unexpectedly.");
        }
    }
    public async Task RunMigrationForRecordIdsAsync(List<string> recordIds, List<UserMapping>? userMappings = null)
    {
        var startTime = DateTime.UtcNow;
        var processedCount = 0;
        var errorCount = 0;

        _logger.LogInformation("🚀 Starting migration at {StartTime}", startTime);

        try
        {
            if (!await _apiService.LoginAsync())
            {
                _logger.LogError("❌ Failed to login to API, aborting migration");
                return;
            }

            var recordsList = _configuration["SharePoint:RecordsListName"] ?? "User_Records";
            var tasksList = _configuration["SharePoint:TasksListName"] ?? "User_Records_Associated_Task";
            var internalTasksList = _configuration["SharePoint:InternalTasksListName"] ?? "Internal_Records_Associated_Task";
            var internalList = _configuration["SharePoint:InternalListName"] ?? "Internal_Records";
            var rowLimit = _configuration.GetValue<int?>("SharePoint:RowLimit") ?? 500;

            
            

            foreach (var recordId in recordIds)
            {
                var id = int.Parse(recordId);
                
                try
                {
                    var record = await _sharePointService.GetRecordByIdAsync(recordsList, id, userMappings);

                    if (record == null)
                    {
                        _logger.LogWarning("⚠️ Record {Id} not found in SharePoint.", id);
                        continue;
                    }

                    if (record.Record_Type != "InternalMemo")
                    {
                        var recordOfficerRecord = await _sharePointService.GetRecordOfficerDetailsByIdAsync(id, userMappings);
                        if (recordOfficerRecord == null)
                        {
                            continue;
                        }
                        record.Reference_Number = recordOfficerRecord.Reference_Number;
                        record.DateOfDispatch = recordOfficerRecord.DateOfDispatch;
                        record.DispatchedBy = recordOfficerRecord.DispatchedBy;
                        record.RecordId = recordOfficerRecord.ID;
                    }

                    await ProcessRecordAsync(record, recordsList, tasksList, internalList, internalTasksList, userMappings);
                    await SaveCheckpointAsync(record.ID);
                    processedCount++;

                    var success = new MigrationSuccess
                    {
                        RecordId = record.ID,
                        RecordsList = recordsList,
                        TasksList = tasksList,
                        Timestamp = DateTime.UtcNow
                    };

                    await LogSuccessToFileAsync(success,
                        $"{recordIds[0]}_{recordIds[recordIds.Count() -1 ]}_success.json");
                }
                catch (Exception ex)
                {
                    errorCount++;
                    _logger.LogError(ex, "❌ Error processing record {RecordId}", id);

                    var payloadJson = (ex is MigrationException mex && mex.Payload != null)
                        ? JsonSerializer.Serialize(mex.Payload, new JsonSerializerOptions { WriteIndented = true })
                        : null;

                    var error = new MigrationError
                    {
                        RecordId = id,
                        RecordsList = recordsList,
                        TasksList = tasksList,
                        InternalList = internalList,
                        InternalTasksList = internalTasksList,
                        ErrorMessage = ex.ToString(),
                        ErrorType = ErrorTypeOf(ex),
                        Payload = payloadJson,
                        Timestamp = DateTime.UtcNow
                    };

                    await LogErrorToFileAsync(error, $"{recordIds[0]}_{recordIds[recordIds.Count() - 1]}_errors.json");
                }
            }

            var endTime = DateTime.UtcNow;
            _logger.LogInformation("✅ Migration completed successfully at {EndTime}. Duration: {Duration} minutes",
                endTime, (endTime - startTime).TotalMinutes);

            // 📌 Save run log with range info
            var runLog = new MigrationRunLog
            {
                StartTime = startTime,
                EndTime = endTime,
                TotalProcessed = processedCount,
                TotalErrors = errorCount,
            };

            await LogRunAsync(runLog, "migration_runs.json");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "💥 Migration failed unexpectedly.");
        }
    }
    public async Task RunMigrationForRecordIdsDocLibraryAsync(
        int startId,
        int endId,
        List<UserMapping>? userMappings = null)
    {
        var incomingTasksList = _configuration["SharePoint:IncomingTasksListName"] ?? "Incoming_Record_Associated_Task";
        var incomingList = _configuration["SharePoint:IncomingListName"] ?? "IncomingRecords";
        var libraryName = "Records";
        var incomingFolderName = "Incoming";

        var overallStart = DateTime.UtcNow;
        _logger.LogInformation("🚀 Starting document library migration for {Library} at {StartTime}", libraryName, overallStart);

        try
        {
            if (!await _apiService.LoginAsync())
            {
                _logger.LogError("❌ Failed to login to API, aborting document library migration");
                return;
            }

            var folderIds = await _sharePointService.GetIncomingFolderIdsAsync(
                libraryName,
                incomingFolderName,
                startId - 1);

            var requestedFolderIds = folderIds
                .Where(id => id >= startId && id <= endId)
                .ToList();

            _logger.LogInformation(
                "📂 Found {Count} folders under {Library}/{IncomingFolder} in ID range {StartId}-{EndId}.",
                requestedFolderIds.Count,
                libraryName,
                incomingFolderName,
                startId,
                endId);

            foreach (var id in requestedFolderIds)
            {
                var folder = await _sharePointService.GetIncomingFolderByIdAsync(
                    libraryName,
                    id,
                    userMappings);

                if (folder == null)
                {
                    _logger.LogWarning("⚠️ Folder item {FolderId} was not found in {Library}.", id, libraryName);
                    continue;
                }

                var normalizedFolderUrl = Uri.UnescapeDataString(folder.FileRef ?? string.Empty)
                    .Replace('\\', '/')
                    .TrimEnd('/');
                var requiredFolderPath = $"/{libraryName}/{incomingFolderName}/";

                if (!normalizedFolderUrl.Contains(requiredFolderPath, StringComparison.OrdinalIgnoreCase))
                {
                    _logger.LogWarning(
                        "⚠️ Skipping item {FolderId} because it is outside {Library}/{IncomingFolder}. FileRef: {FileRef}",
                        id,
                        libraryName,
                        incomingFolderName,
                        folder.FileRef);
                    continue;
                }

                string folderName = folder.FileLeafRef;
                string folderServerRelativeUrl = folder.FileRef;

                var startTime = DateTime.UtcNow;
                _logger.LogInformation("────────── 📌 Start processing Folder: {FolderName} (ID {FolderId}) at {StartTime} ──────────",
                    folderName, id, startTime);

                try
                {
                    // 2) Process this folder (payload + tasks + files)
                    await ProcessExternalRecordAsync(
                        folder,
                        folderName,
                        folderServerRelativeUrl,
                        id,
                        incomingList,
                        incomingTasksList,
                        userMappings: userMappings);

                    // 3) Save checkpoint
                    await SaveExternalIncomingCheckpointAsync(id);

                    var endTime = DateTime.UtcNow;
                    var duration = (endTime - startTime).TotalSeconds;

                    // 4) Log success
                    var success = new MigrationSuccess
                    {
                        RecordId = id,
                        RecordsList = libraryName,
                        TasksList = incomingFolderName,
                        Timestamp = endTime,
                        StartTime = startTime,
                        EndTime = endTime,
                        DurationSeconds = duration
                    };

                    await LogIncomingSuccessToFileAsync(success, $"{libraryName}_doclib_success.json");

                    _logger.LogInformation("✅ Completed Folder {FolderName} (ID {FolderId}) at {EndTime} | Duration: {Duration}s",
                        folderName, id, endTime, duration);
                }
                catch (Exception ex)
                {
                    var endTime = DateTime.UtcNow;
                    var duration = (endTime - startTime).TotalSeconds;
                    var payloadJson = (ex is MigrationException mex && mex.Payload != null)
                       ? JsonSerializer.Serialize(mex.Payload, new JsonSerializerOptions { WriteIndented = true })
                       : null;

                    _logger.LogError(ex, "❌ Error processing folder {FolderName} (ID {FolderId}) at {EndTime} | Duration: {Duration}s",
                        folderName, id, endTime, duration);

                    var error = new MigrationError
                    {
                        RecordId = id,
                        RecordsList = libraryName,
                        TasksList = incomingFolderName,
                        ErrorMessage = ex.ToString(),
                        ErrorType = ErrorTypeOf(ex),
                        Timestamp = endTime,
                        StartTime = startTime,
                        Payload = payloadJson,
                        EndTime = endTime,
                        DurationSeconds = duration,
                        IncomingList = incomingList,
                        IncomingTasksList = incomingTasksList,

                    };

                    await LogIncomingErrorToFileAsync(error, $"{libraryName}_doclib_errors.json");
                }

                _logger.LogInformation("────────────────────────────────────────────────────────────");
            }

            var overallEnd = DateTime.UtcNow;
            _logger.LogInformation("✅ Document library migration finished for {Library} at {EndTime}. Duration: {Duration} minutes",
                libraryName, overallEnd, (overallEnd - overallStart).TotalMinutes);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "💥 Document library migration failed unexpectedly for {Library}", libraryName);
        }
    }

    /// <summary>
    /// User_Records items ready for migration: every item in the list with Status = Sent.
    /// IDs in the *_success.json logs are counted as already migrated and omitted unless includeMigrated is set.
    /// </summary>
    /// <param name="recordType">Optional Record_Type filter (case-insensitive); comma-separate several, e.g. "Internal,Outgoing".</param>
    public async Task<ReadyRecordsResult> GetReadyUserRecordsAsync(bool includeMigrated = false, string? recordType = null)
    {
        var recordsList = _configuration["SharePoint:RecordsListName"] ?? "User_Records";
        var migrated = await LoadMigratedIdsAsync(isIncoming: false);
        var failed = await LoadLatestErrorsAsync(isIncoming: false);

        _logger.LogInformation("[ready] Loading Sent items from {List}", recordsList);
        var sent = await _sharePointService.GetSentRecordSummariesAsync(recordsList);

        // Counts of every Record_Type among the Sent items, so the valid filter values are visible.
        var availableTypes = sent
            .GroupBy(r => RecordTypeKey(r.RecordType))
            .ToDictionary(g => g.Key, g => g.Count());

        var wanted = ParseRecordTypes(recordType);
        if (wanted.Count > 0)
        {
            sent = sent.Where(r => wanted.Contains(RecordTypeKey(r.RecordType))).ToList();
            _logger.LogInformation("[ready] Filtered to Record_Type {Types}: {Count} item(s)", string.Join(",", wanted), sent.Count);
        }

        var result = new ReadyRecordsResult { Source = recordsList, TotalFound = sent.Count, AvailableRecordTypes = availableTypes };
        foreach (var r in sent)
        {
            if (migrated.Contains(r.Id)) { result.AlreadyMigrated++; if (!includeMigrated) continue; }
            else if (failed.TryGetValue(r.Id, out var err)) AddFailed(result, r.Id, err);
            else result.NotMigratedIds.Add(r.Id);
            result.Ids.Add(r.Id);
        }
        return result;
    }

    /// <summary>
    /// Folders under Records/Incoming ready for migration. IDs in Records_doclib_success.json are counted as
    /// already migrated and omitted unless includeMigrated is set.
    /// </summary>
    public async Task<ReadyRecordsResult> GetReadyIncomingRecordsAsync(bool includeMigrated = false)
    {
        var migrated = await LoadMigratedIdsAsync(isIncoming: true);
        var failed = await LoadLatestErrorsAsync(isIncoming: true);

        _logger.LogInformation("[ready] Loading Records/Incoming folders");
        var folders = await _sharePointService.GetIncomingFolderSummariesAsync("Records", "Incoming");

        var result = new ReadyRecordsResult { Source = "Records/Incoming", TotalFound = folders.Count };
        foreach (var f in folders)
        {
            if (migrated.Contains(f.Id)) { result.AlreadyMigrated++; if (!includeMigrated) continue; }
            else if (failed.TryGetValue(f.Id, out var err)) AddFailed(result, f.Id, err);
            else result.NotMigratedIds.Add(f.Id);
            result.Ids.Add(f.Id);
        }
        return result;
    }

    // "Internal, Outgoing" -> {"Internal","Outgoing"} (case-insensitive). Empty set = no filter.
    private static HashSet<string> ParseRecordTypes(string? recordType) =>
        (recordType ?? string.Empty)
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

    private static string RecordTypeKey(string? type) =>
        string.IsNullOrWhiteSpace(type) ? "(none)" : type.Trim();

    private static void AddFailed(ReadyRecordsResult result, int id, MigrationError error)
    {
        var (category, reason) = ClassifyFailure(error);
        result.FailedIds.Add(id);
        result.FailedDetails.Add(new FailedRecordInfo
        {
            Id = id,
            ErrorType = string.IsNullOrEmpty(error.ErrorType) ? MigrationErrorTypes.Migration : error.ErrorType,
            Category = category,
            Reason = reason,
            FailedAt = error.Timestamp
        });
        if (!result.FailedByCategory.TryGetValue(category, out var ids))
            result.FailedByCategory[category] = ids = new List<int>();
        ids.Add(id);
        result.FailureSummary[category] = ids.Count;
    }

    // Best-effort classification from the logged error text. The API's own message (after " | API: ") is preferred
    // because it names the real cause; entries logged before that was captured fall back to the exception text.
    private static (string Category, string Reason) ClassifyFailure(MigrationError error)
    {
        var text = error.ErrorMessage ?? string.Empty;
        var apiAt = text.IndexOf(ApiMarker, StringComparison.Ordinal);
        string reason;
        if (apiAt >= 0)
        {
            reason = text[(apiAt + ApiMarker.Length)..];
            var stack = reason.IndexOf("\n   at ", StringComparison.Ordinal);
            if (stack >= 0) reason = reason[..stack];
        }
        else
        {
            reason = text.Split('\n')[0].Trim();
        }
        reason = reason.Trim();
        if (reason.Length > 300) reason = reason[..300] + "…";

        var t = text.ToLowerInvariant();
        string category;
        if (t.Contains("violates not-null")) category = "MissingRequiredField";
        else if (t.Contains("violates foreign key")) category = "InvalidReference";
        else if (t.Contains("duplicate key") || t.Contains("already exists")) category = "Duplicate";
        else if (t.Contains("fileinfo should not be empty")) category = "NoFiles";
        else if (t.Contains("jwt") || t.Contains(" api: 401") || t.Contains(" api: 403")) category = "Auth";
        else if (t.Contains(" api: 404") || t.Contains("not found")) category = "NotFound";
        else if (t.Contains("minio") || t.Contains("s3exception") || t.Contains("amazons3")) category = "Storage";
        else if (t.Contains("timeout") || t.Contains("task was canceled") || t.Contains("httprequestexception") || t.Contains("connection")) category = "Network";
        else if (t.Contains(" api: 5")) category = "ApiServerError";
        else if (t.Contains(" api: 4")) category = "ApiRejected";
        else if (error.ErrorType == MigrationErrorTypes.FileInfo) category = "FileInfoUnknown";
        else category = "Unknown";

        return (category, reason);
    }

    // Last logged error per id across the matching *_errors.json logs.
    private async Task<Dictionary<int, MigrationError>> LoadLatestErrorsAsync(bool isIncoming)
    {
        var latest = new Dictionary<int, MigrationError>();
        foreach (var e in await ReadLogEntriesAsync<MigrationError>(_configuration["SharePoint:ErrorFilePath"], "*_errors.json", "_doclib_errors.json", isIncoming))
            if (!latest.TryGetValue(e.RecordId, out var prev) || e.Timestamp >= prev.Timestamp)
                latest[e.RecordId] = e;
        return latest;
    }

    // Success logs: Incoming runs write "<library>_doclib_success.json"; User_Records runs write "<start>_<end>_success.json".
    private async Task<HashSet<int>> LoadMigratedIdsAsync(bool isIncoming) =>
        await LoadLoggedIdsAsync(_configuration["SharePoint:SuccessFilePath"], "*_success.json", "_doclib_success.json", isIncoming);

    private async Task<HashSet<int>> LoadLoggedIdsAsync(string? configuredPath, string pattern, string docLibSuffix, bool isIncoming)
    {
        var ids = new HashSet<int>();
        foreach (var e in await ReadLogEntriesAsync<LoggedId>(configuredPath, pattern, docLibSuffix, isIncoming))
            ids.Add(e.RecordId);
        return ids;
    }

    private async Task<List<T>> ReadLogEntriesAsync<T>(string? configuredPath, string pattern, string docLibSuffix, bool isIncoming)
    {
        var entries = new List<T>();
        var basePath = configuredPath ?? AppContext.BaseDirectory;
        if (!Directory.Exists(basePath)) return entries;

        foreach (var file in Directory.EnumerateFiles(basePath, pattern))
        {
            var isDocLib = Path.GetFileName(file).EndsWith(docLibSuffix, StringComparison.OrdinalIgnoreCase);
            if (isDocLib != isIncoming) continue;

            try
            {
                var json = await File.ReadAllTextAsync(file);
                if (string.IsNullOrWhiteSpace(json)) continue;
                entries.AddRange(JsonSerializer.Deserialize<List<T>>(json) ?? new());
            }
            catch (JsonException ex)
            {
                _logger.LogWarning(ex, "⚠️ Could not read log {File}", file);
            }
        }

        return entries;
    }

    private class LoggedId { public int RecordId { get; set; } }

    public async Task RunDocumentLibraryMigrationAsync()
    {
        var incomingTasksList = _configuration["SharePoint:IncomingTasksListName"] ?? "Incoming_Record_Associated_Task";
        var incomingList = _configuration["SharePoint:IncomingListName"] ?? "IncomingRecords";
        var libraryName = "Records";
        var incomingFolderName = "Incoming";

        var overallStart = DateTime.UtcNow;
        _logger.LogInformation("🚀 Starting document library migration for {Library} at {StartTime}", libraryName, overallStart);

        try
        {
            if (!await _apiService.LoginAsync())
            {
                _logger.LogError("❌ Failed to login to API, aborting document library migration");
                return;
            }

            var lastProcessedId = await GetLastExternalIncomingProcessedIdAsync();


            // 1) Get folder IDs inside Incoming
            var folders = await _sharePointService.GetIncomingFolderIdsAsync(libraryName, incomingFolderName, lastProcessedId);
            _logger.LogInformation("📂 Found {Count} folders under {Incoming}", folders.Count, incomingFolderName);

            foreach (var folderId in folders)
            {
                var folder = await _sharePointService.GetIncomingFolderByIdAsync(libraryName, folderId);
                string folderName = folder.FileLeafRef;
                string folderServerRelativeUrl = folder.FileRef;

                var startTime = DateTime.UtcNow;
                _logger.LogInformation("────────── 📌 Start processing Folder: {FolderName} (ID {FolderId}) at {StartTime} ──────────",
                    folderName, folderId, startTime);

                try
                {
                    // 2) Process this folder (payload + tasks + files)
                    await ProcessExternalRecordAsync(folder, folderName, folderServerRelativeUrl, folderId, incomingList, incomingTasksList);

                    // 3) Save checkpoint
                    await SaveExternalIncomingCheckpointAsync(folderId);

                    var endTime = DateTime.UtcNow;
                    var duration = (endTime - startTime).TotalSeconds;

                    // 4) Log success
                    var success = new MigrationSuccess
                    {
                        RecordId = folderId,
                        RecordsList = libraryName,
                        TasksList = incomingFolderName,
                        Timestamp = endTime,
                        StartTime = startTime,
                        EndTime = endTime,
                        DurationSeconds = duration
                    };

                    await LogIncomingSuccessToFileAsync(success, $"{libraryName}_doclib_success.json");

                    _logger.LogInformation("✅ Completed Folder {FolderName} (ID {FolderId}) at {EndTime} | Duration: {Duration}s",
                        folderName, folderId, endTime, duration);
                }
                catch (Exception ex)
                {
                    var endTime = DateTime.UtcNow;
                    var duration = (endTime - startTime).TotalSeconds;
                    var payloadJson = (ex is MigrationException mex && mex.Payload != null)
                       ? JsonSerializer.Serialize(mex.Payload, new JsonSerializerOptions { WriteIndented = true })
                       : null;

                    _logger.LogError(ex, "❌ Error processing folder {FolderName} (ID {FolderId}) at {EndTime} | Duration: {Duration}s",
                        folderName, folderId, endTime, duration);

                    var error = new MigrationError
                    {
                        RecordId = folderId,
                        RecordsList = libraryName,
                        TasksList = incomingFolderName,
                        ErrorMessage = ex.ToString(),
                        ErrorType = ErrorTypeOf(ex),
                        Timestamp = endTime,
                        StartTime = startTime,
                        Payload = payloadJson,
                        EndTime = endTime,
                        DurationSeconds = duration,
                        IncomingList = incomingList,
                        IncomingTasksList = incomingTasksList,

                    };

                    await LogIncomingErrorToFileAsync(error, $"{libraryName}_doclib_errors.json");
                }

                _logger.LogInformation("────────────────────────────────────────────────────────────");
            }

            var overallEnd = DateTime.UtcNow;
            _logger.LogInformation("✅ Document library migration finished for {Library} at {EndTime}. Duration: {Duration} minutes",
                libraryName, overallEnd, (overallEnd - overallStart).TotalMinutes);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "💥 Document library migration failed unexpectedly for {Library}", libraryName);
        }
    }
     public async Task FindAndStoreAllUsersInJsonForDocumentLibary()
    {
        var incomingTasksList = _configuration["SharePoint:IncomingTasksListName"] ?? "Incoming_Record_Associated_Task";
        var incomingList = _configuration["SharePoint:IncomingListName"] ?? "IncomingRecords";
        var libraryName = "Records";
        var incomingFolderName = "Incoming";

        var overallStart = DateTime.UtcNow;
        _logger.LogInformation("🚀 Starting document library migration for {Library} at {StartTime}", libraryName, overallStart);

        try
        {
            if (!await _apiService.LoginAsync())
            {
                _logger.LogError("❌ Failed to login to API, aborting document library migration");
                return;
            }

           
            // 1) Get folder IDs inside Incoming
            var folders = await _sharePointService.GetIncomingFolderIdsAsync(libraryName, incomingFolderName, 0);
            _logger.LogInformation("📂 Found {Count} folders under {Incoming}", folders.Count, incomingFolderName);

            foreach (var folderId in folders)
            {
                var folder = await _sharePointService.GetIncomingFolderByIdAsync(libraryName, folderId);
                string folderName = folder.FileLeafRef;
                string folderServerRelativeUrl = folder.FileRef;

                var startTime = DateTime.UtcNow;
                _logger.LogInformation("────────── 📌 Start processing Folder: {FolderName} (ID {FolderId}) at {StartTime} ──────────",
                    folderName, folderId, startTime);

                try
                {
                    // 2) Process this folder (payload + tasks + files)
                    await ProcessExternalRecordAsync(folder, folderName, folderServerRelativeUrl, folderId, incomingList, incomingTasksList, false);
                 
                }
                catch (Exception ex)
                {
                    var endTime = DateTime.UtcNow;
                    var duration = (endTime - startTime).TotalSeconds;
                    var payloadJson = (ex is MigrationException mex && mex.Payload != null)
                       ? JsonSerializer.Serialize(mex.Payload, new JsonSerializerOptions { WriteIndented = true })
                       : null;

                    _logger.LogError(ex, "❌ Error processing folder {FolderName} (ID {FolderId}) at {EndTime} | Duration: {Duration}s",
                        folderName, folderId, endTime, duration);

                    var error = new MigrationError
                    {
                        RecordId = folderId,
                        RecordsList = libraryName,
                        TasksList = incomingFolderName,
                        ErrorMessage = ex.ToString(),
                        ErrorType = ErrorTypeOf(ex),
                        Timestamp = endTime,
                        StartTime = startTime,
                        Payload = payloadJson,
                        EndTime = endTime,
                        DurationSeconds = duration,
                        IncomingList = incomingList,
                        IncomingTasksList = incomingTasksList,

                    };

                    await LogIncomingErrorToFileAsync(error, $"{libraryName}_doclib_errors.json");
                }

                _logger.LogInformation("────────────────────────────────────────────────────────────");
            }

            var overallEnd = DateTime.UtcNow;
            _logger.LogInformation("✅ Document library migration finished for {Library} at {EndTime}. Duration: {Duration} minutes",
                libraryName, overallEnd, (overallEnd - overallStart).TotalMinutes);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "💥 Document library migration failed unexpectedly for {Library}", libraryName);
        }
    }

    public async Task RetryErrors(string fileName, List<UserMapping>? userMappings = null)
    {

        var errors = await LoadErrorsFromFileAsync(fileName);
        var results = new List<object>();

        if (!await _apiService.LoginAsync())
        {
            _logger.LogError("❌ Failed to login to API, aborting migration");
            return;
        }

        foreach (var error in errors)
        {
            try
            {
                var success = await ProcessRecordByIdAsync(
                    error.RecordsList,
                    error.TasksList,
                    error.InternalList,
                    error.InternalTasksList,
                    error.RecordId,
                    userMappings
                );

                if (success)
                {
                    // ✅ log success back into the same range file
                    var successEntry = new MigrationSuccess
                    {
                        RecordId = error.RecordId,
                        RecordsList = error.RecordsList,
                        TasksList = error.TasksList,
                        Timestamp = DateTime.UtcNow
                    };

                    await LogSuccessToFileAsync(successEntry, fileName.Replace("_errors.json", "_success.json"));

                    results.Add(new { id = error.RecordId, success = true });
                }
                else
                {
                    // ✅ retry failed again → keep it in error file with payload
                    var retryError = new MigrationError
                    {
                        RecordId = error.RecordId,
                        RecordsList = error.RecordsList,
                        TasksList = error.TasksList,
                        InternalList = error.InternalList,
                        InternalTasksList = error.InternalTasksList,
                        ErrorMessage = "Retry failed again",
                        ErrorType = error.ErrorType,
                        Payload = error.Payload, // keep original payload for reference
                        Timestamp = DateTime.UtcNow
                    };

                    await LogErrorToFileAsync(retryError, fileName);

                    results.Add(new { id = error.RecordId, success = false, error = "Retry failed again" });
                }
            }
            catch (Exception ex)
            {
                // ❌ unexpected exception, log with payload if available
                var retryError = new MigrationError
                {
                    RecordId = error.RecordId,
                    RecordsList = error.RecordsList,
                    TasksList = error.TasksList,
                    InternalList = error.InternalList,
                    InternalTasksList = error.InternalTasksList,
                    ErrorMessage = ex.ToString(),
                        ErrorType = ErrorTypeOf(ex),
                    Payload = error.Payload,
                    Timestamp = DateTime.UtcNow
                };

                await LogErrorToFileAsync(retryError, fileName);

                results.Add(new { id = error.RecordId, success = false, error = ex.Message });
            }

           
        }
    }

    /// <summary>
    /// Migrates the given Records/Incoming folder IDs (no range scan), applying the user mappings.
    /// </summary>
    public async Task RunIncomingForFolderIdsAsync(List<int> folderIds, List<UserMapping>? userMappings = null)
    {
        var incomingTasksList = _configuration["SharePoint:IncomingTasksListName"] ?? "Incoming_Record_Associated_Task";
        var incomingList = _configuration["SharePoint:IncomingListName"] ?? "IncomingRecords";

        if (!await _apiService.LoginAsync())
        {
            _logger.LogError("❌ Failed to login to API, aborting incoming migration");
            return;
        }

        foreach (var id in folderIds.Distinct())
            await ProcessFolderById("Records", id, incomingList, incomingTasksList, userMappings);
    }

    public async Task<bool> ProcessFolderById(string libraryName, int folderId, string incomingList, string incomingTasksList, List<UserMapping>? userMappings = null)
    {
        var folder = await _sharePointService.GetIncomingFolderByIdAsync(libraryName, folderId, userMappings);
        string folderName = folder.FileLeafRef;
        string folderServerRelativeUrl = folder.FileRef;

        var startTime = DateTime.UtcNow;
        _logger.LogInformation("────────── 📌 Start processing Folder: {FolderName} (ID {FolderId}) at {StartTime} ──────────",
            folderName, folderId, startTime);

        try
        {
            // 2) Process this folder (payload + tasks + files)
            await ProcessExternalRecordAsync(folder, folderName, folderServerRelativeUrl, folderId, incomingList, incomingTasksList, userMappings: userMappings);

            // 3) Save checkpoint
            await SaveExternalIncomingCheckpointAsync(folderId);

            var endTime = DateTime.UtcNow;
            var duration = (endTime - startTime).TotalSeconds;

            // 4) Log success
            var success = new MigrationSuccess
            {
                RecordId = folderId,
                RecordsList = libraryName,
                Timestamp = endTime,
                StartTime = startTime,
                EndTime = endTime,
                DurationSeconds = duration
            };

            await LogIncomingSuccessToFileAsync(success, $"{libraryName}_doclib_success.json");

            _logger.LogInformation("✅ Completed Folder {FolderName} (ID {FolderId}) at {EndTime} | Duration: {Duration}s",
                folderName, folderId, endTime, duration);
            return true;
        }
        catch (Exception ex)
        {
            var endTime = DateTime.UtcNow;
            var duration = (endTime - startTime).TotalSeconds;
            var payloadJson = (ex is MigrationException mex && mex.Payload != null)
               ? JsonSerializer.Serialize(mex.Payload, new JsonSerializerOptions { WriteIndented = true })
               : null;

            _logger.LogError(ex, "❌ Error processing folder {FolderName} (ID {FolderId}) at {EndTime} | Duration: {Duration}s",
                folderName, folderId, endTime, duration);

            var error = new MigrationError
            {
                RecordId = folderId,
                RecordsList = libraryName,
                ErrorMessage = ex.ToString(),
                        ErrorType = ErrorTypeOf(ex),
                Timestamp = endTime,
                StartTime = startTime,
                Payload = payloadJson,
                EndTime = endTime,
                DurationSeconds = duration,
                IncomingList = incomingList,
                IncomingTasksList = incomingTasksList,

            };

            await LogIncomingErrorToFileAsync(error, $"{libraryName}_doclib_errors.json");
            return false;

        }

    }
    public async Task RetryErrorsDocLib(string fileName, List<UserMapping>? userMappings = null)
    {

        var errors = await LoadErrorsFromFileAsync(fileName);
        var results = new List<object>();

        if (!await _apiService.LoginAsync())
        {
            _logger.LogError("❌ Failed to login to API, aborting migration");
            return;
        }

        foreach (var error in errors)
        {
            try
            {
                var success = await ProcessFolderById(
                    "Records",
                    error.RecordId,
                    error.IncomingList,
                    error.IncomingTasksList,
                    userMappings
                );

                if (success)
                {
                    // ✅ log success back into the same range file
                    var successEntry = new MigrationSuccess
                    {
                        RecordId = error.RecordId,
                        RecordsList = error.RecordsList,
                        TasksList = error.TasksList,
                        Timestamp = DateTime.UtcNow
                    };

                    await LogSuccessToFileAsync(successEntry, fileName.Replace("_errors.json", "_success.json"));

                    results.Add(new { id = error.RecordId, success = true });
                }
                else
                {
                    // ✅ retry failed again → keep it in error file with payload
                    var retryError = new MigrationError
                    {
                        RecordId = error.RecordId,
                        RecordsList = error.RecordsList,
                        TasksList = error.TasksList,
                        InternalList = error.InternalList,
                        InternalTasksList = error.InternalTasksList,
                        ErrorMessage = "Retry failed again",
                        ErrorType = error.ErrorType,
                        Payload = error.Payload, // keep original payload for reference
                        Timestamp = DateTime.UtcNow
                    };

                    await LogErrorToFileAsync(retryError, fileName);

                    results.Add(new { id = error.RecordId, success = false, error = "Retry failed again" });
                }
            }
            catch (Exception ex)
            {
                // ❌ unexpected exception, log with payload if available
                var retryError = new MigrationError
                {
                    RecordId = error.RecordId,
                    RecordsList = error.RecordsList,
                    TasksList = error.TasksList,
                    InternalList = error.InternalList,
                    InternalTasksList = error.InternalTasksList,
                    ErrorMessage = ex.ToString(),
                        ErrorType = ErrorTypeOf(ex),
                    Payload = error.Payload,
                    Timestamp = DateTime.UtcNow
                };

                await LogErrorToFileAsync(retryError, fileName);

                results.Add(new { id = error.RecordId, success = false, error = ex.Message });
            }

           
        }
    }

    public async Task<bool> ProcessRecordByIdAsync(
    string recordsList,
    string tasksList,
    string internalList,
    string internalTasksList,
    int recordId,
    List<UserMapping>? userMappings = null)
    {
        SharePointRecord? record = null;

        try
        {
            // ✅ Fetch record
            record = await _sharePointService.GetRecordByIdAsync(recordsList, recordId, userMappings);

            if (record == null)
            {
                _logger.LogWarning("⚠️ Record {RecordId} not found in list {List}", recordId, recordsList);
                return false;
            }

            // ✅ Process record (attachments, payload build, save to DB, etc.)
            await ProcessRecordAsync(record, recordsList, tasksList, internalList, internalTasksList, userMappings);

            // ✅ Save checkpoint after successful processing
            await SaveCheckpointAsync(record.ID);

            _logger.LogInformation("✅ Successfully processed record {RecordId}", record.ID);
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "❌ Error processing record {RecordId}", record?.ID ?? recordId);
            return false;
        }
    }

    private async Task ProcessRecordAsync(
    SharePointRecord record,
    string recordsList,
    string tasksList,
    string internalList,
    string internalTasksList,
    List<UserMapping> userMappings = null,
    bool process = true)
    {
        _logger.LogInformation("▶ Processing record {Id} ({Title})", record.ID, record.Title);

        // 1) Fetch related tasks
        var tasks = await _sharePointService.GetTasksForRecordAsync(tasksList, record.ID, 500, userMappings)
                    ?? new List<SharePointTask>();
        _logger.LogInformation("  → Found {Count} tasks for record {Id}", tasks.Count, record.ID);

        var incomingTasks = await _sharePointService.GetParentWithIncomingTasksAsync(
            internalList, internalTasksList, record.RecordId, userMappings
        );
        _logger.LogInformation("  → Found {Count} incoming tasks for record {Id}", incomingTasks.Count, record.ID);

        if(process)
        {
            // 2) Build payload & send migration data
            var payload = MapToMigrationPayload(record, tasks, incomingTasks);
            var ok = await _apiService.MigrateRecordAsync(payload);

            if (!ok)
            {
                throw new MigrationException(
                    $"❌ Failed to migrate Record ID: {record.ID}{ApiDetail()}",
                    payload
                );
            }

            // 3) Check MinIO for files
            var (fileInfo, attachments) = await _minioService.GetRecordFilesAsync(record.ID);

            if (fileInfo != null || attachments.Count > 0)
            {
                _logger.LogInformation("📂 Using existing files from MinIO for record {Id}", record.ID);
            }
            else
            {
                _logger.LogInformation("⬆️ No files found in MinIO for record {Id}, uploading from SharePoint...", record.ID);
                (fileInfo, attachments) = await UploadAttachmentsAsync(recordsList, record.ID, record.Title);
                _logger.LogInformation("  → Uploaded {Count} attachments for record {Id}", attachments.Count, record.ID);
            }

            // 4) Update file info in API (only if there’s something to update)
            var fileUpdate = new FileInfoMigrationDto
            {
                Id = record.ID.ToString(),
                FileInfo = fileInfo,
                Attachments = attachments
            };

            if (fileUpdate.FileInfo != null || fileUpdate.Attachments.Count > 0)
            {
                var updateOk = await _apiService.UpdateRecordFilesAsync(fileUpdate);

                if (!updateOk)
                {
                    throw new MigrationException(
                        $"❌ Failed to update file info for Record ID: {record.ID}{ApiDetail()}",
                        fileUpdate,
                        errorType: MigrationErrorTypes.FileInfo
                    );
                }

                _logger.LogInformation("📎 File info updated for Record ID: {RecordId}", record.ID);
            }

            _logger.LogInformation("✅ Fully migrated Record ID: {RecordId}", record.ID);
        }
    }

    private async Task ProcessExternalRecordAsync(
    SharePointFolder record,
    string folderName,
    string folderServerRelativeUrl,
    int folderId,
    string incomingList,
    string incomingTasksList,
    bool process = true,
    List<UserMapping>? userMappings = null)
    {
        _logger.LogInformation("▶ [Step 1] Building payload for folder {FolderName} (ID {Id})", folderName, folderId);

        // 1) Get incoming tasks
        var incomingTasks = await _sharePointService.GetParentWithExternalIncomingTasksAsync(
            incomingList, incomingTasksList, record.ID, userMappings);

        _logger.LogInformation("   → Found {Count} incoming tasks for folder {Id}", incomingTasks.Count, folderId);

        if(process)
        {
            // 2) Map payload
            var payload = MapToIncomingRecordPayload(record, incomingTasks);
            _logger.LogInformation("   → Payload built for folder {Id} with ReferenceNumber: {RefNum}", folderId, payload.referenceNumber);

            // 3) Send payload to API
            var ok = await _apiService.MigrateExternalRecordAsync(payload);
            if (!ok)
            {
                throw new MigrationException($"❌ Failed to migrate folder {folderName}{ApiDetail()}", payload);
            }
            _logger.LogInformation("   ✅ Payload migration succeeded for folder {Id}", folderId);

            // 4) Get / Upload files
            _logger.LogInformation("▶ [Step 2] Processing files for folder {FolderName} (ID {Id})", folderName, folderId);
            var (fileInfo, attachments) = await _minioService.GetRecordFilesFromFolderAsync(folderId);

            if (fileInfo != null || attachments.Count > 0)
            {
                _logger.LogInformation("   📂 Using existing files from MinIO for folder {Id}", folderId);
            }
            else
            {
                _logger.LogInformation("   ⬆️ Uploading files from SharePoint for folder {Id}", folderId);
                (fileInfo, attachments) = await UploadFilesAsync(folderServerRelativeUrl, folderId, record.RecordFileName);
                _logger.LogInformation("   → Uploaded {Count} attachments for folder {Id}", attachments.Count, folderId);
            }

            // 5) Update file info in API
            _logger.LogInformation("▶ [Step 3] Updating file info for folder {FolderName} (ID {Id})", folderName, folderId);

            var fileUpdate = new FileInfoMigrationDto
            {
                Id = folderId.ToString(),
                FileInfo = fileInfo,
                Attachments = attachments
            };

            var updateOk = await _apiService.UpdateRecordFilesAsync(fileUpdate);
            if (!updateOk)
            {
                throw new MigrationException($"❌ Failed to update file info for Record ID: {record.ID}{ApiDetail()}", fileUpdate, errorType: MigrationErrorTypes.FileInfo);
            }

            _logger.LogInformation("   📎 File info updated for folder {Id}", folderId);
        }
        _logger.LogInformation("✅ Migration complete for folder {FolderName} (ID {Id})", folderName, folderId);
    }



    // ===== Upload helpers =====
    private async Task<(MinIOFileInfo fileInfo, List<MinIOFileInfo> attachments)>
    UploadAttachmentsAsync(string recordsList, int recordId, string recordTitle)
    {
        var files = await _sharePointService.GetItemAttachmentsAsync(recordsList, recordId);

        if (files == null || files.Count == 0)
            return (null, new List<MinIOFileInfo>());

        return await _minioService.UploadListItemAttachmentsAsync(recordId, files, SafeTitle(recordTitle));
    }
    
    private async Task<(MinIOFileInfo fileInfo, List<MinIOFileInfo> attachments)>
    UploadFilesAsync(string folderServerRelativeUrl, int recordId ,  string recordTitle)
    {
        var files = await _sharePointService.GetFilesInFolderAsync(folderServerRelativeUrl);

        if (files == null || files.Count == 0)
            return (null, new List<MinIOFileInfo>());

        return await _minioService.UploadFilesAsync(recordId, files, SafeTitle(recordTitle));
    }

    private const string ApiMarker = " | API: ";

    // Appends the target API's last error body so the error log says why the API refused.
    private string ApiDetail() =>
        string.IsNullOrWhiteSpace(_apiService.LastError) ? string.Empty : ApiMarker + _apiService.LastError;

    private static string ErrorTypeOf(Exception ex) =>
        (ex as MigrationException)?.ErrorType ?? MigrationErrorTypes.Migration;

    public Task LogErrorToFileAsync(MigrationError error, string fileName) =>
        AppendToJsonListFileAsync(ErrorFilePath(fileName), error);

    public Task LogIncomingErrorToFileAsync(MigrationError error, string fileName) =>
        AppendToJsonListFileAsync(ErrorFilePath(fileName), error);

    public Task LogSuccessToFileAsync(MigrationSuccess success, string fileName) =>
        AppendToJsonListFileAsync(SuccessFilePath(fileName), success);

    private Task LogIncomingSuccessToFileAsync(MigrationSuccess success, string fileName) =>
        AppendToJsonListFileAsync(SuccessFilePath(fileName), success);

    public async Task<List<MigrationError>> LoadErrorsFromFileAsync(string fileName)
    {
        await _resultFileLock.WaitAsync();
        try
        {
            return await ReadJsonListFileAsync<MigrationError>(ErrorFilePath(fileName));
        }
        finally
        {
            _resultFileLock.Release();
        }
    }

    private string ErrorFilePath(string fileName) =>
        Path.Combine(_configuration["SharePoint:ErrorFilePath"] ?? AppContext.BaseDirectory, fileName);

    private string SuccessFilePath(string fileName) =>
        Path.Combine(_configuration["SharePoint:SuccessFilePath"] ?? AppContext.BaseDirectory, fileName);

    // The error/success/run files are JSON arrays rewritten in full on every append. Rewriting in place
    // left a cut-off file whenever the app was stopped mid-write, so writes now go to a temp file that
    // replaces the original in one step. One lock serialises them: the service is a singleton shared by requests.
    private static readonly SemaphoreSlim _resultFileLock = new(1, 1);
    private static readonly JsonSerializerOptions _resultFileJson = new() { WriteIndented = true };

    private async Task AppendToJsonListFileAsync<T>(string path, T item)
    {
        await _resultFileLock.WaitAsync();
        try
        {
            var items = await ReadJsonListFileAsync<T>(path);
            items.Add(item);
            await WriteJsonListFileAsync(path, items);
        }
        finally
        {
            _resultFileLock.Release();
        }
    }

    // Caller must hold _resultFileLock.
    private async Task<List<T>> ReadJsonListFileAsync<T>(string path)
    {
        if (!File.Exists(path))
            return new List<T>();

        var json = await File.ReadAllTextAsync(path);
        if (string.IsNullOrWhiteSpace(json))
            return new List<T>();

        try
        {
            return JsonSerializer.Deserialize<List<T>>(json) ?? new List<T>();
        }
        catch (JsonException ex)
        {
            // Damaged file (e.g. cut off by an earlier crash): keep every complete entry, save the original
            // next to it, and rewrite the file so it loads normally from now on.
            var recovered = SalvageJsonList<T>(System.Text.Encoding.UTF8.GetBytes(json));
            var backupPath = $"{path}.corrupt-{DateTime.Now:yyyyMMdd-HHmmss}";
            File.Copy(path, backupPath, overwrite: true);
            await WriteJsonListFileAsync(path, recovered);

            _logger.LogWarning(ex,
                "⚠️ {Path} was damaged. Recovered {Count} complete entries; the damaged original was saved as {Backup}.",
                path, recovered.Count, backupPath);
            return recovered;
        }
    }

    private static async Task WriteJsonListFileAsync<T>(string path, List<T> items)
    {
        var tempPath = path + ".tmp";
        await File.WriteAllTextAsync(tempPath, JsonSerializer.Serialize(items, _resultFileJson));
        File.Move(tempPath, path, overwrite: true);
    }

    // Returns the complete elements of a JSON array, stopping at the first one that is cut off or malformed.
    private static List<T> SalvageJsonList<T>(byte[] utf8)
    {
        var items = new List<T>();
        var reader = new Utf8JsonReader(utf8);
        try
        {
            if (!reader.Read() || reader.TokenType != JsonTokenType.StartArray)
                return items;

            while (reader.Read() && reader.TokenType != JsonTokenType.EndArray)
            {
                var start = (int)reader.TokenStartIndex;
                reader.Skip(); // throws if this element never closes
                var element = JsonSerializer.Deserialize<T>(utf8.AsSpan(start, (int)reader.BytesConsumed - start));
                if (element != null)
                    items.Add(element);
            }
        }
        catch (JsonException)
        {
            // The damaged element and anything after it are dropped here; they remain in the backup file.
        }
        return items;
    }


    private static string SafeTitle(string title) =>
        string.IsNullOrWhiteSpace(title) ? "record" : title.Trim();

    // Records with a blank subject fail to migrate, so send a placeholder instead.
    private static string SubjectOrDefault(string? subject) =>
        string.IsNullOrWhiteSpace(subject) ? "Empty Subject" : subject;
    private static string MapRecordType(string type)
    {
        if (string.IsNullOrWhiteSpace(type)) return "internal"; // default

        switch (type.Trim().ToLowerInvariant())
        {
            case "internal": return "internal";
            case "outgoing": return "external";
            case "internalmemo": return "internal_memo";
            default: return type.ToLowerInvariant();
        }
    }

    private static string MapStatusKey(string status)
    {
        if (string.IsNullOrWhiteSpace(status)) return "draft"; // default

        switch (status.Trim())
        {
            case "Draft": return "Draft";
            case "Forwarded": return "Submitted";
            case "Approved": return "Approved";
            case "Accepted": return "Submitted";
            case "Adjustment": return "Adjustment";
            case "Rejected": return "Rejected";
            case "Returned": return "Adjustment";
            case "Sent": return "Dispatched";
            default: return status.ToLowerInvariant();
        }
    }
    private static string MapIncomingStatusKey(string status)
    {
        if (string.IsNullOrWhiteSpace(status)) return "draft"; // default
        switch (status.Trim())
        {
            case "In Progress": return "draft";
            case "Draft": return "draft";
            case "Accepted": return "accepted";
            case "Responded": return "accepted";
            case "Rejected": return "rejected";
            case "Reassigned": return "re_assigned";
            case "Returned": return "rejected";
            default: return status.ToLowerInvariant();
        }
    }

    // ===== Mapping =====
    private MigrationPayload MapToMigrationPayload(
        SharePointRecord record,
        List<SharePointTask> tasks,
        List<SharePointTask> incomingTasks)
    {
        var taskInstances = tasks.Select(t => new OutgoingInstance
        {
            //currentStepId = "nazifjadwa",
            currentStepId = t.AssignedTo ?? "",
            dispatchedBy = t.Author ?? "",
            //dispatchedBy = "wederyeleshwadwa",
            dispatchedAt = t.Created,
            taskStatus = t.TaskStatus ?? "",
            comment = t.Reassign_x0020_Remark ?? ""
        }).ToList();

        var incomingInstances = incomingTasks.Select(t => new IncomingInstance
        {
            id = t.ID.ToString(),
            assignedAt = t.Created,
            //assignerUsername = "fanuelsmart",
            assignerUsername = t.Author,
            parentAssignmentId = t.Parent_task_id == 0 ? null : t.Parent_task_id.ToString(),
            //assigneePositionKey = "nazifjadwa",
            assigneePositionKey = t.AssignedTo,
            //assigneeUsername = "nazifjadwa",
            assigneeUsername = t.AssignedTo,
            status = MapIncomingStatusKey(t.TaskStatus ?? ""),
            comment = t.Reassign_x0020_Remark ?? ""
        }).ToList();

        if (taskInstances.Any())
        {
            taskInstances.Last().comment = tasks.Last().Reassign_x0020_Remark ?? "";
        }
        if (incomingInstances.Any())
        {
            incomingInstances.Last().comment = incomingInstances.Last().comment ?? "";
        }

        var contentItem = new ContentItem
        {
            body = record.Original_Content ?? "",
            subject = SubjectOrDefault(record.Subject),
            date = record.GregorianDate.HasValue
                ? DateTime.SpecifyKind(record.GregorianDate.Value, DateTimeKind.Utc)
                : null,

            // ✅ map receiver/cc fields (split CSVs into lists)
            nonSmartOfficeReceiverCCs = ToListFromCsv(record.CC_Organization_Name),
            nonSmartOfficeReceivers = ToListFromCsv(record.Recipient_Organization_Name),

            sincerelyText = "", 
            prefix = record.Prefix ?? "",
            suffix = record.Suffix ?? "",
            prefixCC = record.PrefixCC ?? "",
            suffixCC = record.SuffixCC ?? "",

            // ✅ user/group fields (split string into list if multiple usernames)
            //receiverEmployees = ["nazifjadwa"],
            receiverEmployees = record.Receiving_Users,
            receiverPositions = record.Receiving_Department,
            //ccPosition = ["የኮሚኒኬሽን ዳይሬክቶሬት ዳይሬክተር"],
            ccPosition = record.CC_To,

            // ✅ tasks
            outgoingInstances = taskInstances,
            outgoingHistories = new List<object>(),
        };

        var unitKey = _configuration["MinIO:UnitKey"];
        var payload =  new MigrationPayload
        {
            id = record.ID,
            createdAt = record.Created,
            submittedDate = record.Modified,
            dispatchedDate = record.DateOfDispatch,
            //dispatchedBy = "wederyeleshwadwa",
            dispatchedBy = record.DispatchedBy,
            letterNumber = record.Reference_Number ?? "",
            referenceNumber = record.Reference_Number ?? "",
            //initiatorId = "kumelagadwa",
            initiatorId = record.Author ?? "",
            preferredLanguage = record.RecordLanguage ?? "",
            recordTypeKey = MapRecordType(record.Record_Type),
            statusKey = MapStatusKey(record.Status),
            unitId = new List<int>(),
            unitKey = unitKey,
            incomingInstances = incomingInstances,
            metadata = new MetadataBlock
            {
                organizationName = record.Sending_Organization_Name ?? "",
                deliveredBy = record.Sending_Department ?? ""
            },
            content = new List<ContentItem> { contentItem },
        };

        if (tasks.Count > 0)
        {
            //payload.completedBy = "wederyeleshwadwa";
            payload.completedBy = tasks.Last().Author;
            payload.completedDate = tasks.Last().Created;
        }

        return payload;
    }

    private Task LogRunAsync(MigrationRunLog runLog, string fileName) =>
        AppendToJsonListFileAsync(
            Path.Combine(_configuration["SharePoint:RunLogPath"] ?? AppContext.BaseDirectory, fileName),
            runLog);


    private IncomingRecordMigrationPayload MapToIncomingRecordPayload(
    SharePointFolder folder,
    
    List<SharePointTask> incomingTasks)
    {
        var incomingInstances = incomingTasks.Select(t => new IncomingInstance
        {
            id = t.ID.ToString(),
            assignedAt = t.Created,
            //assignerUsername = "fanuelsmart",
            assignerUsername = t.Author,
            parentAssignmentId = t.Parent_x0020_ID == 0 ? null : t.Parent_x0020_ID.ToString(),
            //assigneePositionKey = "nazifjadwa",
            assigneePositionKey = t.AssignedTo,
            //assigneeUsername = "nazifjadwa",
            assigneeUsername = t.AssignedTo,
            status = MapIncomingStatusKey(t.TaskStatus ?? ""),
        }).ToList();

        if (incomingInstances.Any())
        {
            incomingInstances.Last().comment = incomingInstances.Last().comment ?? "";
        }

        var unitKey = _configuration["MinIO:UnitKey"];


        return new IncomingRecordMigrationPayload
        {
            referenceNumber = folder.Reference_Number,
            letterNumber = folder.RecordFileName, // assuming this is the letter/file ref
            unitKey = unitKey,

            metadata = new Metadata
            {
                organizationName = new LocalizedText
                {
                    en = folder.Sending_Organization_Name ?? "",
                    am = folder.Sending_Organization_Name ?? ""
                },
                deliveredBy = new LocalizedText
                {
                    en = folder.Delivery_Personnel_Name ?? "",
                    am = folder.Delivery_Personnel_Name ?? ""
                },
                uploadedBy = new LocalizedText
                {
                    en = folder.Author ?? "",
                    am = folder.Author ?? ""
                },
            },

            createdAt = DateTime.TryParse(folder.Incoming_Record_Date, out var created)
                ? created.ToString("o") // ISO 8601
                : DateTime.UtcNow.ToString("o"),

            comment = "", // can be populated later

            dispatchedDate = DateTime.TryParse(folder.Incoming_Record_Date, out var dispatchDate)
                ? dispatchDate
                : DateTime.UtcNow,

            dispatchedBy = folder.Author ?? "",

            subject = SubjectOrDefault(folder.Subject),
            body = "",
            id = folder.ID.ToString() ?? "",

            incomingInstances = incomingInstances ?? new List<IncomingInstance>()
        };
    }



    // ===== Checkpointing =====
    private async Task<int> GetLastProcessedIdAsync()
    {
        var path = _configuration["SharePoint:CheckpointFile"];
        if (string.IsNullOrWhiteSpace(path))
        {
            var root = _configuration["SharePoint:OutputPath"] ?? Path.Combine(AppContext.BaseDirectory, "output");
            Directory.CreateDirectory(root);
            path = Path.Combine(root, "checkpoint.log");
        }

        if (!File.Exists(path)) return 0;
        var lastLine = (await File.ReadAllLinesAsync(path)).LastOrDefault();
        return int.TryParse(lastLine, out var id) ? id : 0;
    }
    private async Task<int> GetLastExternalIncomingProcessedIdAsync()
    {
        var path = _configuration["SharePoint:ExternalIncomingCheckpointFile"];
        if (string.IsNullOrWhiteSpace(path))
        {
            var root = _configuration["SharePoint:OutputPath"] ?? Path.Combine(AppContext.BaseDirectory, "output");
            Directory.CreateDirectory(root);
            path = Path.Combine(root, "checkpoint.log");
        }

        if (!File.Exists(path)) return 0;
        var lastLine = (await File.ReadAllLinesAsync(path)).LastOrDefault();
        return int.TryParse(lastLine, out var id) ? id : 0;
    }

    private async Task SaveCheckpointAsync(int recordId)
    {
        var path = _configuration["SharePoint:CheckpointFile"];
        if (string.IsNullOrWhiteSpace(path))
        {
            var root = _configuration["SharePoint:OutputPath"] ?? Path.Combine(AppContext.BaseDirectory, "output");
            Directory.CreateDirectory(root);
            path = Path.Combine(root, "checkpoint.log");
        }
        await File.AppendAllTextAsync(path, recordId + Environment.NewLine);
    } 
    private async Task SaveExternalIncomingCheckpointAsync(int recordId)
    {
        var path = _configuration["SharePoint:ExternalIncomingCheckpointFile"];
        if (string.IsNullOrWhiteSpace(path))
        {
            var root = _configuration["SharePoint:OutputPath"] ?? Path.Combine(AppContext.BaseDirectory, "output");
            Directory.CreateDirectory(root);
            path = Path.Combine(root, "checkpoint.log");
        }
        await File.AppendAllTextAsync(path, recordId + Environment.NewLine);
    }

    private static List<string> ToListFromCsv(string? csv)
    {
        if (string.IsNullOrWhiteSpace(csv))
            return new List<string>();

        // ✅ assign back, remove "[]"
        csv = csv.Replace("[]", string.Empty);

        return csv.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
                  .Select(x => x.Trim())
                  .Where(x => !string.IsNullOrWhiteSpace(x) && x != "[]") // ✅ extra guard
                  .ToList();
    }


    private static List<string> SplitUsers(string? users)
    {
        if (string.IsNullOrWhiteSpace(users)) return new List<string>();
        return users.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(u => u.Trim())
                    .Where(u => !string.IsNullOrWhiteSpace(u))
                    .ToList();
    }

}

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

    public async Task RunMigrationAsync(int? startId = null, int? endId = null, List<UserMapping> userMappings = null)
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
                    var record = await _sharePointService.GetRecordByIdAsync(recordsList, id, userMappings);

                    if (record == null)
                    {
                        _logger.LogWarning("⚠️ Record {Id} not found in SharePoint.", id);
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
    public async Task RunMigrationForRecordIdsAsync(List<string> recordIds)
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
                    var record = await _sharePointService.GetRecordByIdAsync(recordsList, id);

                    if (record == null)
                    {
                        _logger.LogWarning("⚠️ Record {Id} not found in SharePoint.", id);
                        continue;
                    }

                    if (record.Record_Type != "InternalMemo")
                    {
                        var recordOfficerRecord = await _sharePointService.GetRecordOfficerDetailsByIdAsync(id);
                        if (recordOfficerRecord == null)
                        {
                            continue;
                        }
                        record.Reference_Number = recordOfficerRecord.Reference_Number;
                        record.DateOfDispatch = recordOfficerRecord.DateOfDispatch;
                        record.DispatchedBy = recordOfficerRecord.DispatchedBy;
                        record.RecordId = recordOfficerRecord.ID;
                    }

                    await ProcessRecordAsync(record, recordsList, tasksList, internalList, internalTasksList);
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
    public async Task<ReadyRecordsResult> GetReadyUserRecordsAsync(bool includeMigrated = false)
    {
        var recordsList = _configuration["SharePoint:RecordsListName"] ?? "User_Records";
        var migrated = await LoadMigratedIdsAsync(isIncoming: false);
        var failed = await LoadFailedIdsAsync(isIncoming: false);

        _logger.LogInformation("[ready] Loading Sent items from {List}", recordsList);
        var sent = await _sharePointService.GetSentRecordSummariesAsync(recordsList);

        var result = new ReadyRecordsResult { Source = recordsList, TotalFound = sent.Count };
        foreach (var r in sent)
        {
            if (migrated.Contains(r.Id)) { result.AlreadyMigrated++; if (!includeMigrated) continue; }
            else if (failed.Contains(r.Id)) result.FailedIds.Add(r.Id);
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
        var failed = await LoadFailedIdsAsync(isIncoming: true);

        _logger.LogInformation("[ready] Loading Records/Incoming folders");
        var folders = await _sharePointService.GetIncomingFolderSummariesAsync("Records", "Incoming");

        var result = new ReadyRecordsResult { Source = "Records/Incoming", TotalFound = folders.Count };
        foreach (var f in folders)
        {
            if (migrated.Contains(f.Id)) { result.AlreadyMigrated++; if (!includeMigrated) continue; }
            else if (failed.Contains(f.Id)) result.FailedIds.Add(f.Id);
            else result.NotMigratedIds.Add(f.Id);
            result.Ids.Add(f.Id);
        }
        return result;
    }

    // Success logs: Incoming runs write "<library>_doclib_success.json"; User_Records runs write "<start>_<end>_success.json".
    private async Task<HashSet<int>> LoadMigratedIdsAsync(bool isIncoming) =>
        await LoadLoggedIdsAsync(_configuration["SharePoint:SuccessFilePath"], "*_success.json", "_doclib_success.json", isIncoming);

    // Error logs follow the same naming with "_errors.json". An id that later succeeded is excluded by the caller
    // (migrated is checked first), so retried-and-fixed records don't show as failed.
    private async Task<HashSet<int>> LoadFailedIdsAsync(bool isIncoming) =>
        await LoadLoggedIdsAsync(_configuration["SharePoint:ErrorFilePath"], "*_errors.json", "_doclib_errors.json", isIncoming);

    private async Task<HashSet<int>> LoadLoggedIdsAsync(string? configuredPath, string pattern, string docLibSuffix, bool isIncoming)
    {
        var ids = new HashSet<int>();
        var basePath = configuredPath ?? AppContext.BaseDirectory;
        if (!Directory.Exists(basePath)) return ids;

        foreach (var file in Directory.EnumerateFiles(basePath, pattern))
        {
            var isDocLib = Path.GetFileName(file).EndsWith(docLibSuffix, StringComparison.OrdinalIgnoreCase);
            if (isDocLib != isIncoming) continue;

            try
            {
                var json = await File.ReadAllTextAsync(file);
                if (string.IsNullOrWhiteSpace(json)) continue;
                foreach (var e in JsonSerializer.Deserialize<List<LoggedId>>(json) ?? new())
                    ids.Add(e.RecordId);
            }
            catch (JsonException ex)
            {
                _logger.LogWarning(ex, "⚠️ Could not read log {File}", file);
            }
        }

        return ids;
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

    public async Task RetryErrors(string fileName)
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
                    error.RecordId
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
                    Payload = error.Payload,
                    Timestamp = DateTime.UtcNow
                };

                await LogErrorToFileAsync(retryError, fileName);

                results.Add(new { id = error.RecordId, success = false, error = ex.Message });
            }

           
        }
    }

    public async Task<bool> ProcessFolderById(string libraryName, int folderId, string incomingList, string incomingTasksList )
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
    public async Task RetryErrorsDocLib(string fileName)
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
                    error.IncomingTasksList
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
    int recordId)
    {
        SharePointRecord? record = null;

        try
        {
            // ✅ Fetch record
            record = await _sharePointService.GetRecordByIdAsync(recordsList, recordId);

            if (record == null)
            {
                _logger.LogWarning("⚠️ Record {RecordId} not found in list {List}", recordId, recordsList);
                return false;
            }

            // ✅ Process record (attachments, payload build, save to DB, etc.)
            await ProcessRecordAsync(record, recordsList, tasksList, internalList, internalTasksList);

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
                    $"❌ Failed to migrate Record ID: {record.ID}",
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
                        $"❌ Failed to update file info for Record ID: {record.ID}",
                        fileUpdate
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
                throw new MigrationException($"❌ Failed to migrate folder {folderName}", payload);
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
                throw new MigrationException($"❌ Failed to update file info for Record ID: {record.ID}", fileUpdate);
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

    public async Task LogErrorToFileAsync(MigrationError error, string fileName)
    {
        // Base folder path from config
        string basePath = _configuration["SharePoint:ErrorFilePath"] ?? AppContext.BaseDirectory;

        // Final path includes the fileName
        string errorFilePath = Path.Combine(basePath, fileName);

        List<MigrationError> existing = new();

        // Read existing errors if file exists
        if (File.Exists(errorFilePath))
        {
            var json = await File.ReadAllTextAsync(errorFilePath);
            if (!string.IsNullOrWhiteSpace(json))
            {
                existing = JsonSerializer.Deserialize<List<MigrationError>>(json) ?? new List<MigrationError>();
            }
        }

        // Append new error
        existing.Add(error);

        // Write back to file (formatted for readability)
        var updatedJson = JsonSerializer.Serialize(existing, new JsonSerializerOptions { WriteIndented = true });
        await File.WriteAllTextAsync(errorFilePath, updatedJson);
    }
    public async Task LogIncomingErrorToFileAsync(MigrationError error, string fileName)
    {
        // Base folder path from config
        string basePath = _configuration["SharePoint:ErrorFilePath"] ?? AppContext.BaseDirectory;

        // Final path includes the fileName
        string errorFilePath = Path.Combine(basePath, fileName);

        List<MigrationError> existing = new();

        // Read existing errors if file exists
        if (File.Exists(errorFilePath))
        {
            var json = await File.ReadAllTextAsync(errorFilePath);
            if (!string.IsNullOrWhiteSpace(json))
            {
                existing = JsonSerializer.Deserialize<List<MigrationError>>(json) ?? new List<MigrationError>();
            }
        }

        // Append new error
        existing.Add(error);

        // Write back to file (formatted for readability)
        var updatedJson = JsonSerializer.Serialize(existing, new JsonSerializerOptions { WriteIndented = true });
        await File.WriteAllTextAsync(errorFilePath, updatedJson);
    }

    public async Task LogSuccessToFileAsync(MigrationSuccess success, string fileName)
    {
        string basePath = _configuration["SharePoint:SuccessFilePath"] ?? AppContext.BaseDirectory;
        string successFilePath = Path.Combine(basePath, fileName);

        List<MigrationSuccess> existing = new();

        if (File.Exists(successFilePath))
        {
            var json = await File.ReadAllTextAsync(successFilePath);
            if (!string.IsNullOrWhiteSpace(json))
            {
                existing = JsonSerializer.Deserialize<List<MigrationSuccess>>(json) ?? new List<MigrationSuccess>();
            }
        }

        existing.Add(success);

        var updatedJson = JsonSerializer.Serialize(existing, new JsonSerializerOptions { WriteIndented = true });
        await File.WriteAllTextAsync(successFilePath, updatedJson);
    }
    private async Task LogIncomingSuccessToFileAsync(MigrationSuccess success, string fileName)
    {
        string basePath = _configuration["SharePoint:SuccessFilePath"] ?? AppContext.BaseDirectory;
        string successFilePath = Path.Combine(basePath, fileName);

        List<MigrationSuccess> existing = new();

        if (File.Exists(successFilePath))
        {
            var json = await File.ReadAllTextAsync(successFilePath);
            if (!string.IsNullOrWhiteSpace(json))
            {
                existing = JsonSerializer.Deserialize<List<MigrationSuccess>>(json) ?? new List<MigrationSuccess>();
            }
        }

        existing.Add(success);

        var updatedJson = JsonSerializer.Serialize(existing, new JsonSerializerOptions { WriteIndented = true });
        await File.WriteAllTextAsync(successFilePath, updatedJson);
    }

    public async Task<List<MigrationError>> LoadErrorsFromFileAsync(string fileName)
    {
        string basePath = _configuration["SharePoint:ErrorFilePath"] ?? AppContext.BaseDirectory;
        string errorFilePath = Path.Combine(basePath, fileName);

        if (!File.Exists(errorFilePath))
            return new List<MigrationError>();

        var json = await File.ReadAllTextAsync(errorFilePath);
        if (string.IsNullOrWhiteSpace(json))
            return new List<MigrationError>();

        return JsonSerializer.Deserialize<List<MigrationError>>(json) ?? new List<MigrationError>();
    }


    private static string SafeTitle(string title) =>
        string.IsNullOrWhiteSpace(title) ? "record" : title.Trim();
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
            subject = record.Subject ?? "",
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

    private async Task LogRunAsync(MigrationRunLog runLog, string fileName)
    {
        string basePath = _configuration["SharePoint:RunLogPath"] ?? AppContext.BaseDirectory;
        string filePath = Path.Combine(basePath, fileName);

        List<MigrationRunLog> existing = new();

        if (File.Exists(filePath))
        {
            var json = await File.ReadAllTextAsync(filePath);
            if (!string.IsNullOrWhiteSpace(json))
            {
                existing = JsonSerializer.Deserialize<List<MigrationRunLog>>(json) ?? new List<MigrationRunLog>();
            }
        }

        existing.Add(runLog);

        var updatedJson = JsonSerializer.Serialize(existing, new JsonSerializerOptions { WriteIndented = true });
        await File.WriteAllTextAsync(filePath, updatedJson);
    }


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

            subject = folder.Subject ?? "",
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

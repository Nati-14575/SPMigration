using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using SPMigration.Models;

public class SharePointService
{
    private readonly HttpClient _client;
    private readonly string _siteUrl;
    private readonly string _absUrl;
    private readonly string _userFile;
    private readonly HashSet<string> _uniqueUsers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    public SharePointService(IConfiguration config)
    {
        _siteUrl = config["SharePoint:SiteUrl"];
        _absUrl = config["SharePoint:AbsUrl"];
        _userFile = config["SharePoint:UniqueUsers"];

        var handler = new HttpClientHandler
        {
            Credentials = new NetworkCredential(
                config["SharePoint:Username"],
                config["SharePoint:Password"],
                config["SharePoint:Domain"]
            )
        };

        // Equivalent to Node.js rejectUnauthorized: false. Keep validation enabled
        // unless the setting is explicitly disabled for a trusted migration network.
        var rejectUnauthorized = config.GetValue("SharePoint:RejectUnauthorized", true);
        if (!rejectUnauthorized)
        {
            handler.ServerCertificateCustomValidationCallback =
                HttpClientHandler.DangerousAcceptAnyServerCertificateValidator;
        }

        _client = new HttpClient(handler)
        {
            Timeout = TimeSpan.FromMinutes(10)
        };
        LoadExistingUsers();
    }

    private void TrackUser(object input)
    {
        if (input == null) return;

        if (input is string userName)
        {
            AddUser(userName);
        }
        else if (input is IEnumerable<string> users)
        {
            foreach (var user in users)
            {
                AddUser(user);
            }
        }
    }

    private void AddUser(string userName)
    {
        if (!string.IsNullOrWhiteSpace(userName) && _uniqueUsers.Add(userName))
        {
            SaveUsers(new MigratedUser
            {
                UserName = userName
            });
        }
    }


    private void SaveUsers(MigratedUser newUser)
    {
        if (string.IsNullOrWhiteSpace(_userFile))
        {
            throw new InvalidOperationException("SharePoint:UniqueUsers must be configured with a file path.");
        }

        var directory = Path.GetDirectoryName(_userFile);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        List<MigratedUser> users;
        if (File.Exists(_userFile))
        {
            var json = File.ReadAllText(_userFile);
            users = JsonSerializer.Deserialize<List<MigratedUser>>(json) ?? new List<MigratedUser>();
        }
        else
        {
            users = new List<MigratedUser>();
        }

        users.Add(newUser);
        var updatedJson = JsonSerializer.Serialize(users, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(_userFile, updatedJson);
    }

    private void LoadExistingUsers()
    {
        if (File.Exists(_userFile))
        {
            var json = File.ReadAllText(_userFile);
            var users = JsonSerializer.Deserialize<List<MigratedUser>>(json) ?? new List<MigratedUser>();

            foreach (var user in users)
            {
                if (!string.IsNullOrWhiteSpace(user.UserName))
                    _uniqueUsers.Add(user.UserName);
            }
        }
    }

    // ✅ Get new digest (needed for POSTs)
    private async Task<string> GetRequestDigestAsync()
    {
        var url = $"{_siteUrl}/_api/contextinfo";
        var request = new HttpRequestMessage(HttpMethod.Post, url);
        request.Headers.Add("Accept", "application/json;odata=verbose");

        var response = await _client.SendAsync(request);
        var result = await response.Content.ReadAsStringAsync();

        using var doc = JsonDocument.Parse(result);
        return doc.RootElement
                  .GetProperty("d")
                  .GetProperty("GetContextWebInformation")
                  .GetProperty("FormDigestValue")
                  .GetString();
    }

    // ✅ Generic RLDS call
    private async Task<JsonDocument> InvokeRLDSAsync(string listTitle, string viewXml, string nextHref = null)
    {
        if (!string.IsNullOrEmpty(nextHref))
        {
            // follow paging link
            var resp = await _client.GetAsync(nextHref);
            var json = await resp.Content.ReadAsStringAsync();
            return JsonDocument.Parse(json);
        }

        var digest = await GetRequestDigestAsync();
        var url = $"{_siteUrl}/_api/web/lists/getbytitle('{listTitle}')/RenderListDataAsStream";

        var payload = new
        {
            parameters = new
            {
                RenderOptions = 2,
                ViewXml = viewXml,
                AddRequiredFields = true
            }
        };

        var req = new HttpRequestMessage(HttpMethod.Post, url);
        req.Headers.Add("Accept", "application/json;odata=nometadata");
        req.Headers.Add("X-RequestDigest", digest);
        req.Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json");

        var response = await _client.SendAsync(req);
        var content = await response.Content.ReadAsStringAsync();

        return JsonDocument.Parse(content);
    }

    // ✅ Get records with RLDS
    public async Task<List<int>> GetRecordIdsAsync(string listTitle, int lastProcessedId, int rowLimit, List<UserMapping> userMappings = null)
    {
        var digest = await GetRequestDigestAsync();

        var whereClause = lastProcessedId > 0
     ? $@"
      <And>
        <Lt>
          <FieldRef Name='ID' />
          <Value Type='Counter'>{lastProcessedId}</Value>
        </Lt>
        <Eq>
          <FieldRef Name='Status' />
          <Value Type='Text'>Sent</Value>
        </Eq>
      </And>"
     : @"
        <Eq>
          <FieldRef Name='Status' />
          <Value Type='Text'>Sent</Value>
        </Eq>";

        var viewXml = $@"
<View>
  <Query>
    <Where>
      {whereClause}
    </Where>
    <OrderBy>
      <FieldRef Name='ID' Ascending='False' />
    </OrderBy>
  </Query>
  <ViewFields>
    <FieldRef Name='ID' />
  </ViewFields>
  <RowLimit Paged='TRUE'>{rowLimit}</RowLimit>
</View>";


        var payload = new
        {
            parameters = new { ViewXml = viewXml, RenderOptions = 2 }
        };

        var req = new HttpRequestMessage(HttpMethod.Post,
            $"{_siteUrl}/_api/web/lists/getbytitle('{listTitle}')/RenderListDataAsStream");
        req.Headers.Add("Accept", "application/json;odata=verbose");
        req.Headers.Add("X-RequestDigest", digest);
        req.Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json");

        var res = await _client.SendAsync(req);
        res.EnsureSuccessStatusCode();
        var raw = await res.Content.ReadAsStringAsync();

        using var doc = JsonDocument.Parse(raw);
        var rows = doc.RootElement.GetProperty("Row");

        return rows.EnumerateArray()
                   .Select(r => int.TryParse(r.GetProperty("ID").GetString(), out var id) ? id : 0)
                   .Where(id => id > 0)
                   .ToList();
    }
    public async Task<SharePointRecord?> GetRecordByIdAsync(string listTitle, int id, List<UserMapping> userMappings = null)
    {
        // Build query with expand for user fields
        var url = $"{_siteUrl}/_api/web/lists/getbytitle('{listTitle}')/items({id})" +
                  "?$select=*," +
                  "Author/Id,Author/Title,Author/Name," +
                  "Editor/Id,Editor/Title,Editor/Name," +
                  "Receiving_Users/Id,Receiving_Users/Title,Receiving_Users/Name," +
                  "Receiving_Department/Id,Receiving_Department/Title,Receiving_Department/Name," +
                  "CC_To/Id,CC_To/Title,CC_To/Name" +
                  "&$expand=Author,Editor,Receiving_Users,Receiving_Department,CC_To";

        var req = new HttpRequestMessage(HttpMethod.Get, url);
        req.Headers.Add("Accept", "application/json;odata=verbose");

        var res = await _client.SendAsync(req);
        if (!res.IsSuccessStatusCode) return null;

        var raw = await res.Content.ReadAsStringAsync();
        using var doc = JsonDocument.Parse(raw);

        var d = doc.RootElement.GetProperty("d");

        // ✅ Only proceed if Status == "Sent"
        var status = d.TryGetProperty("Status", out var st) ? st.GetString() : null;
        if (!string.Equals(status, "Sent", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        var recordType = d.TryGetProperty("Record_Type", out var rt) ? rt.GetString() : null;
        var newRecord = new SharePointRecord
        {
            ID = d.GetProperty("ID").GetInt32(),
            Title = d.TryGetProperty("Title", out var title) ? title.GetString() : null,
            Subject = d.TryGetProperty("Subject", out var subj) ? subj.GetString() : null,
            Original_Content = d.TryGetProperty("Original_Content", out var oc) ? oc.GetString() : null,
            FileLeafRef = d.TryGetProperty("FileLeafRef", out var fl) ? fl.GetString() : null,
            Reference_Number = d.TryGetProperty("Reference_Number", out var rn) ? rn.GetString() : null,
            Record_Type = recordType,
            Recipient_Organization_Name = recordType == "Internal" ? "" : d.TryGetProperty("Receiver", out var ron) ? ron.GetString() : null,
            CC_Organization_Name = recordType == "Internal" ? "" : d.TryGetProperty("CCReceiver", out var con) ? con.GetString() : null,
            Sending_Organization_Name = d.TryGetProperty("Sending_Organization_Name", out var son) ? son.GetString() : null,
            Status = status,
            GregorianDate = d.TryGetProperty("GregorianDate", out var gd) && gd.ValueKind == JsonValueKind.String
                ? DateTime.Parse(gd.GetString()) : null,
            Created = d.TryGetProperty("Created", out var cr) && cr.ValueKind == JsonValueKind.String
                ? DateTime.Parse(cr.GetString()) : DateTime.MinValue,
            Modified = d.TryGetProperty("Modified", out var md) && md.ValueKind == JsonValueKind.String
                ? DateTime.Parse(md.GetString()) : DateTime.MinValue,

            // ✅ User fields
            Author = d.GetUserFieldSafe("Author", userMappings),
            Editor = d.GetUserFieldSafe("Editor", userMappings),
            Receiving_Users = d.ResolveUsersOrGroups("Receiving_Users", userMappings),
            Receiving_Department = d.ResolveUsersOrGroups("Receiving_Department", userMappings),
            CC_To = d.ResolveUsersOrGroups("CC_To", userMappings),
            Sending_Department = d.TryGetProperty("Sending_Department", out var sd) ? sd.GetString() : null
        };
        TrackUser(newRecord.Author);
        TrackUser(newRecord.Editor);
        TrackUser(newRecord.Receiving_Department);
        TrackUser(newRecord.Receiving_Users);
        TrackUser(newRecord.CC_To);
        return newRecord;
    }

    public async Task<SharePointRecord?> GetRecordOfficerDetailsByIdAsync(int id, List<UserMapping> userMappings = null)
    {
        var url = $"{_siteUrl}/_api/web/lists/getbytitle('Records')/items?$filter=RecordId eq {id}" +
                  "&$select=ID,Reference_Number,Date_Of_Dispatch,Author/Id,Author/Title,Author/Name&$expand=Author";

        var req = new HttpRequestMessage(HttpMethod.Get, url);
        req.Headers.Add("Accept", "application/json;odata=verbose");

        var res = await _client.SendAsync(req);
        if (!res.IsSuccessStatusCode) return null;

        var raw = await res.Content.ReadAsStringAsync();
        using var doc = JsonDocument.Parse(raw);

        var results = doc.RootElement.GetProperty("d").GetProperty("results");
        if (results.GetArrayLength() == 0) return null;

        var d = results[0];
        var newRecord = new SharePointRecord
        {
            ID = d.GetProperty("ID").GetInt32(),
            Reference_Number = d.TryGetProperty("Reference_Number", out var rn) ? rn.GetString() : null,
            DispatchedBy = d.GetUserFieldSafe("Author", userMappings),
            DateOfDispatch = d.TryGetProperty("Date_Of_Dispatch", out var cr) && cr.ValueKind == JsonValueKind.String
                ? DateTime.Parse(cr.GetString()) : DateTime.MinValue

        };

        TrackUser(newRecord.DispatchedBy);
        return newRecord;
    }


    // ✅ Get tasks for record with RLDS
    public async Task<List<SharePointTask>> GetTasksForRecordAsync(string listTitle, int recordId, int rowLimit = 500, List<UserMapping> userMappings = null )
    {
        var url =
            $"{_siteUrl}/_api/web/lists/getbytitle('{listTitle}')/items?" +
            $"$select=Id,Parent_x0020_ID,Created,Modified,Body,TaskStatus,Reassign_x0020_Remark," +
            $"Author/Id,Author/Title,Author/Name," +
            $"AssignedTo/Id,AssignedTo/Title,AssignedTo/Name" +
            $"&$expand=Author,AssignedTo" +
            $"&$filter=Parent_x0020_ID eq {recordId}" +
            $"&$orderby=Id asc" +
            $"&$top={rowLimit}";

        var req = new HttpRequestMessage(HttpMethod.Get, url);
        req.Headers.Add("Accept", "application/json;odata=verbose");

        var res = await _client.SendAsync(req);
        res.EnsureSuccessStatusCode();

        var raw = await res.Content.ReadAsStringAsync();
        var tasks = new List<SharePointTask>();

        using var doc = JsonDocument.Parse(raw);
        if (doc.RootElement.TryGetProperty("d", out var d) &&
            d.TryGetProperty("results", out var results))
        {
            foreach (var row in results.EnumerateArray())
            {
                var newTask = new SharePointTask
                {
                    ID = row.GetProperty("Id").GetInt32(),
                    Parent_x0020_ID = row.GetProperty("Parent_x0020_ID").GetInt32(),
                    Created = row.GetProperty("Created").GetDateTime(),
                    Modified = row.GetProperty("Modified").GetDateTime(),
                    TaskStatus = row.TryGetProperty("TaskStatus", out var status) ? status.GetString() ?? "" : "",
                    Reassign_x0020_Remark = row.TryGetProperty("Reassign_x0020_Remark", out var remark) ? remark.GetString() ?? "" : "",

                    // 🔥 Reuse your existing helpers
                    Author = row.GetUserFieldSafe("Author", userMappings),
                    AssignedTo = row.GetUserFieldSafe("AssignedTo", userMappings)
                };
                tasks.Add(newTask);

                TrackUser(newTask.AssignedTo);
                TrackUser(newTask.Author);
            }
        }

        return tasks;
    }
    public async Task<List<SharePointTask>> GetParentTasksForRecordAsync(string listTitle, int recordId, int rowLimit = 500, List<UserMapping> userMappings = null)
    {
        var url =
            $"{_siteUrl}/_api/web/lists/getbytitle('{listTitle}')/items?" +
            $"$select=Id,RecordId,Created,Modified," +
            $"Author/Id,Author/Title,Author/Name," +
            $"ReceivingDepartment/Id,ReceivingDepartment/Title,ReceivingDepartment/Name,Record_Type" +
            $"&$expand=Author,ReceivingDepartment" +
            $"&$filter=RecordId eq {recordId} and Record_Type eq 'Record'" +
            $"&$orderby=Id asc" +
            $"&$top={rowLimit}";

        var req = new HttpRequestMessage(HttpMethod.Get, url);
        req.Headers.Add("Accept", "application/json;odata=verbose");

        var res = await _client.SendAsync(req);
        res.EnsureSuccessStatusCode();

        var raw = await res.Content.ReadAsStringAsync();
        var tasks = new List<SharePointTask>();

        using var doc = JsonDocument.Parse(raw);
        if (doc.RootElement.TryGetProperty("d", out var d) &&
            d.TryGetProperty("results", out var results))
        {
            foreach (var row in results.EnumerateArray())
            {
                var newTask = new SharePointTask
                {
                    ID = row.GetProperty("Id").GetInt32(),
                    Created = row.GetProperty("Created").GetDateTime(),
                    Modified = row.GetProperty("Modified").GetDateTime(),
                    Author = row.GetUserFieldSafe("Author", userMappings),
                    AssignedTo = row.ResolveUsersOrGroups("ReceivingDepartment", userMappings)[0]
                };
                tasks.Add(newTask);

                TrackUser(newTask.AssignedTo);
                TrackUser(newTask.Author);
            }
        }

        return tasks;
    }
    public async Task<List<SharePointTask>> GetParentTasksForIncomingRecordAsync(string listTitle, int recordId, int rowLimit = 500, List<UserMapping> userMappings = null)
    {
        var url =
            $"{_siteUrl}/_api/web/lists/getbytitle('{listTitle}')/items?" +
            $"$select=Id,RecordId,Created,Modified," +
            $"Author/Id,Author/Title,Author/Name," +
            $"AssignedTo/Id,AssignedTo/Title,AssignedTo/Name" +
            $"&$expand=Author,AssignedTo" +
            $"&$filter=RecordId eq {recordId}" +
            $"&$orderby=Id asc" +
            $"&$top={rowLimit}";

        var req = new HttpRequestMessage(HttpMethod.Get, url);
        req.Headers.Add("Accept", "application/json;odata=verbose");

        var res = await _client.SendAsync(req);
        res.EnsureSuccessStatusCode();

        var raw = await res.Content.ReadAsStringAsync();
        var tasks = new List<SharePointTask>();

        using var doc = JsonDocument.Parse(raw);
        if (doc.RootElement.TryGetProperty("d", out var d) &&
            d.TryGetProperty("results", out var results))
        {
            foreach (var row in results.EnumerateArray())
            {
                var newTask = new SharePointTask
                {
                    ID = row.GetProperty("Id").GetInt32(),
                    Created = row.GetProperty("Created").GetDateTime(),
                    Modified = row.GetProperty("Modified").GetDateTime(),
                    Author = row.GetUserFieldSafe("Author", userMappings),
                    AssignedTo = row.ResolveUsersOrGroups("AssignedTo", userMappings)[0]
                };
                tasks.Add(newTask);
                TrackUser(newTask.AssignedTo);
                TrackUser(newTask.Author);
            }
        }

        return tasks;
    }

    public async Task<List<SharePointTask>> GetIncomingTasksForRecordAsync(string listTitle, int recordId, int rowLimit = 500, List<UserMapping> userMappings = null )
    {
        var url =
            $"{_siteUrl}/_api/web/lists/getbytitle('{listTitle}')/items?" +
            $"$select=Id,ParentId,RecordID,Created,Modified,Body,TaskStatus,Reassign_x0020_Remark,Parent_task_idId," +
            $"Author/Id,Author/Title,Author/Name," +
            $"AssignedTo/Id,AssignedTo/Title,AssignedTo/Name" +
            $"&$expand=Author,AssignedTo" +
            $"&$filter=RecordID eq {recordId}" +
            $"&$orderby=Id asc" +
            $"&$top={rowLimit}";

        var req = new HttpRequestMessage(HttpMethod.Get, url);
        req.Headers.Add("Accept", "application/json;odata=verbose");

        var res = await _client.SendAsync(req);
        res.EnsureSuccessStatusCode();

        var raw = await res.Content.ReadAsStringAsync();
        var tasks = new List<SharePointTask>();

        using var doc = JsonDocument.Parse(raw);
        if (doc.RootElement.TryGetProperty("d", out var d) &&
            d.TryGetProperty("results", out var results))
        {
            foreach (var row in results.EnumerateArray())
            {
                var assignedUsers = row.ResolveUsersOrGroups("AssignedTo", userMappings);

                if (assignedUsers == null || !assignedUsers.Any())
                {
                    // ❌ skip tasks without AssignedTo
                    continue;
                }

                // Handle parent task id (string or number)
                int parentTaskId = 0;
                if (row.TryGetProperty("Parent_task_idId", out var ptid))
                {
                    if (ptid.ValueKind == JsonValueKind.String && int.TryParse(ptid.GetString(), out var parsed))
                    {
                        parentTaskId = parsed;
                    }
                    else if (ptid.ValueKind == JsonValueKind.Number)
                    {
                        parentTaskId = ptid.GetInt32();
                    }
                }
                int Id = row.GetProperty("Id").GetInt32();
                foreach (var assigned in assignedUsers)
                {
                    var newTask = new SharePointTask
                    {
                        ID = Id,
                        Created = row.GetProperty("Created").GetDateTime(),
                        Modified = row.GetProperty("Modified").GetDateTime(),
                        TaskStatus = row.TryGetProperty("TaskStatus", out var status) ? status.GetString() ?? "" : "",
                        Reassign_x0020_Remark = row.TryGetProperty("Reassign_x0020_Remark", out var remark) ? remark.GetString() ?? "" : "",
                        Author = row.GetUserFieldSafe("Author", userMappings),
                        AssignedTo = assigned,
                        Parent_task_id = (parentTaskId == Id) ? 0 : parentTaskId
                    }; 
                    tasks.Add(newTask);
                    TrackUser(newTask.AssignedTo);
                    TrackUser(newTask.Author);
                }
            }


        }

        return tasks;
    }

    public async Task<List<SharePointTask>> GetParentWithIncomingTasksAsync(
    string listTitle, string subTaskListTitle, int recordId, List<UserMapping> userMappings = null )
    {
        var subs = await GetIncomingTasksForRecordAsync(subTaskListTitle, recordId, 500, userMappings);

        return subs;
    }
    public async Task<List<SharePointTask>> GetParentWithExternalIncomingTasksAsync(
    string listTitle, string subTaskListTitle, int recordId, List<UserMapping> userMappings = null)
    {
        var subs = await GetIncomingTasksForRecordAsync(subTaskListTitle, recordId, 500, userMappings);


        return subs;
    }


    public async Task<List<(string FileName, byte[] Bytes, string contentType)>> GetItemAttachmentsAsync(string listTitle, int itemId)
    {
        var url = $"{_siteUrl}/_api/web/lists/getbytitle('{listTitle}')/items({itemId})/AttachmentFiles";
        var req = new HttpRequestMessage(HttpMethod.Get, url);
        req.Headers.Add("Accept", "application/json;odata=verbose");

        var res = await _client.SendAsync(req);
        res.EnsureSuccessStatusCode();
        var raw = await res.Content.ReadAsStringAsync();

        var files = new List<(string FileName, byte[] Bytes, string contentType)>();

        using var doc = JsonDocument.Parse(raw);
        var results = doc.RootElement
                         .GetProperty("d")
                         .GetProperty("results");

        foreach (var f in results.EnumerateArray())
        {
            var fileName = f.GetProperty("FileName").GetString();
            var fileUrl = f.GetProperty("ServerRelativeUrl").GetString();

            // Download file bytes
            var fileReq = new HttpRequestMessage(HttpMethod.Get, $"{_absUrl}{fileUrl}");
            var fileRes = await _client.SendAsync(fileReq);
            fileRes.EnsureSuccessStatusCode();

            var bytes = await fileRes.Content.ReadAsByteArrayAsync();
            var contentType = fileRes.Content.Headers.ContentType?.MediaType ?? "application/octet-stream";

            files.Add((fileName, bytes, contentType));
        }


        return files;
    }

    public async Task<List<int>> GetIncomingFolderIdsAsync(
     string libraryName,
     string incomingFolderName = "Incoming",
     int? lastProcessedId = null)
    {
        var ids = new List<int>();


        var filter = lastProcessedId.HasValue
            ? $"$filter=ListItemAllFields/Id gt {lastProcessedId.Value}&"
            : string.Empty;

        var endpoint =
            $"{_siteUrl}/_api/web/GetFolderByServerRelativeUrl('{libraryName}/{incomingFolderName}')/Folders" +
            $"?{filter}$select=ListItemAllFields/Id&$expand=ListItemAllFields&$orderBy=ListItemAllFields/Id asc&$top=5000";

        using var request = new HttpRequestMessage(HttpMethod.Get, endpoint);
        request.Headers.Add("Accept", "application/json;odata=nometadata");

        var response = await _client.SendAsync(request);

        response.EnsureSuccessStatusCode();

        var json = await response.Content.ReadAsStringAsync();
        using var doc = JsonDocument.Parse(json);

        if (doc.RootElement.TryGetProperty("value", out var folders))
        {
            foreach (var folder in folders.EnumerateArray())
            {
                if (!folder.TryGetProperty("ListItemAllFields", out var d))
                    continue;

                if (d.TryGetProperty("Id", out var id))
                {
                    ids.Add(id.GetInt32());
                }
            }
        }

        return ids;
    }

    public async Task<SharePointFolder?> GetIncomingFolderByIdAsync(string libraryName, int itemId, List<UserMapping> userMappings = null)
    {
        var endpoint =
            $"{_siteUrl}/_api/web/lists/getbytitle('{libraryName}')/items({itemId})" +
            "?$select=Id,RecordId,Reference_Number,RecordFileName,Sending_Department,FileRef,FileLeafRef," +
            "Sending_Organization_Name,Recipient_Organization_Name,Delivery_Personnel_Name," +
            "Status,Subject,Created,Modified,GregorianDate,Date_Of_Dispatch,RecordType,Author/Id,Author/Title,Author/Name&" +
            "$expand=Author";

        using var request = new HttpRequestMessage(HttpMethod.Get, endpoint);
        request.Headers.Add("Accept", "application/json;odata=nometadata");

        var response = await _client.SendAsync(request);

        // SharePoint list item IDs are not contiguous. Deleted items leave gaps,
        // so a range migration must treat 404 as "item not found" rather than
        // aborting the entire range.
        if (response.StatusCode == HttpStatusCode.NotFound)
            return null;

        response.EnsureSuccessStatusCode();

        var json = await response.Content.ReadAsStringAsync();
        using var doc = JsonDocument.Parse(json);

        var d = doc.RootElement;

        if (d.ValueKind == JsonValueKind.Undefined || d.ValueKind == JsonValueKind.Null)
            return null;
        var newFolder = new SharePointFolder
        {
            ID = d.TryGetProperty("Id", out var id) ? id.GetInt32() : 0,
            Subject = d.TryGetProperty("Subject", out var subj) ? GetStringSafe(subj) : null,
            FileRef = d.TryGetProperty("FileRef", out var fr) ? GetStringSafe(fr) : null,
            FileLeafRef = d.TryGetProperty("FileLeafRef", out var flr) ? GetStringSafe(flr) : null,
            Sending_Organization_Name = d.TryGetProperty("Sending_Organization_Name", out var son) ? GetStringSafe(son) : null,
            RecordFileName = d.TryGetProperty("RecordFileName", out var rfn) ? GetStringSafe(rfn) : null,
            Incoming_Record_Date = d.TryGetProperty("Created", out var cr) ? GetStringSafe(cr) : null,
            Reference_Number = d.TryGetProperty("Reference_Number", out var rn) ? GetStringSafe(rn) : null,
            Delivery_Personnel_Name = d.TryGetProperty("Delivery_Personnel_Name", out var dpn) ? GetStringSafe(dpn) : null,
            RecordNumber = d.TryGetProperty("RecordId", out var recn) ? GetStringSafe(recn) : null,
            Author = d.GetUserFieldSafe("Author", userMappings)
        };
        TrackUser(newFolder.Author);
        return newFolder;

    }


    public async Task<List<(string FileName, byte[] Bytes, string ContentType)>>
    GetFilesInFolderAsync(string folderServerRelativeUrl)
    {
        var filesResult = new List<(string FileName, byte[] Bytes, string ContentType)>();

        // 1) Get all files metadata under the given folder
        var endpoint = $"{_siteUrl}/_api/web/GetFolderByServerRelativeUrl('{folderServerRelativeUrl}')/Files";

        using var request = new HttpRequestMessage(HttpMethod.Get, endpoint);
        request.Headers.Add("Accept", "application/json;odata=nometadata");

        var response = await _client.SendAsync(request);
        response.EnsureSuccessStatusCode();

        var json = await response.Content.ReadAsStringAsync();
        using var doc = JsonDocument.Parse(json);

        if (doc.RootElement.TryGetProperty("value", out var files))
        {
            foreach (var file in files.EnumerateArray())
            {
                var name = file.GetProperty("Name").GetString();
                var serverRelativeUrl = file.GetProperty("ServerRelativeUrl").GetString();
               

                if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(serverRelativeUrl))
                    continue;

                // 2) Download the file bytes
                var fileEndpoint = $"{_siteUrl}/_api/web/GetFileByServerRelativeUrl('{serverRelativeUrl}')/$value";

                using var fileRequest = new HttpRequestMessage(HttpMethod.Get, fileEndpoint);
                var fileResponse = await _client.SendAsync(fileRequest);
                fileResponse.EnsureSuccessStatusCode();
                var contentType = fileResponse.Content.Headers.ContentType?.MediaType ?? "application/octet-stream";
                var fileBytes = await fileResponse.Content.ReadAsByteArrayAsync();

                filesResult.Add((name, fileBytes, contentType ?? "application/octet-stream"));
            }
        }

        return filesResult;
    }

    private static string GetStringSafe(JsonElement element)
    {
        return element.ValueKind switch
        {
            JsonValueKind.String => element.GetString(),
            JsonValueKind.Number => element.TryGetInt64(out var num) ? num.ToString() : element.ToString(),
            JsonValueKind.Null => null,
            _ => element.ToString()
        };
    }


}

public static class JsonExtensions
{
    public static string? GetStringSafe(this JsonElement element, string propertyName)
    {
        if (!element.TryGetProperty(propertyName, out var prop))
            return null;

        return prop.ValueKind switch
        {
            JsonValueKind.String => prop.GetString(),
            JsonValueKind.Number => prop.ToString(),
            JsonValueKind.True or JsonValueKind.False => prop.GetBoolean().ToString(),
            JsonValueKind.Object => prop.ToString(),
            JsonValueKind.Array => string.Join(";", prop.EnumerateArray().Select(x => x.GetString())),
            _ => null
        };
    }

    public static DateTime? GetDateSafe(this JsonElement element, string propertyName)
    {
        if (!element.TryGetProperty(propertyName, out var prop)) return null;
        if (prop.ValueKind != JsonValueKind.String) return null;

        return DateTime.TryParse(prop.GetString(), out var dt) ? dt : null;
    }

    public static int GetIntSafe(this JsonElement element, string propertyName)
    {
        if (!element.TryGetProperty(propertyName, out var prop)) return 0;
        if (prop.ValueKind == JsonValueKind.String && int.TryParse(prop.GetString(), out var val)) return val;
        if (prop.ValueKind == JsonValueKind.Number && prop.TryGetInt32(out var num)) return num;
        return 0;
    }

    public static string GetUserFieldSafe(this JsonElement element, string propertyName, List<UserMapping> userMappings = null)
    {
        if (!element.TryGetProperty(propertyName, out var prop))
            return "*";

        if (prop.ValueKind == JsonValueKind.Object)
            return ResolveUserOrGroup(prop, userMappings);

        if (prop.ValueKind == JsonValueKind.Array)
        {
            var users = prop.EnumerateArray()
                            .Select(item => ResolveUserOrGroup(item, userMappings))
                            .Where(u => !string.IsNullOrWhiteSpace(u));
            return string.Join(",", users);
        }

        if (prop.ValueKind == JsonValueKind.String)
        {
            var userName = prop.GetString();

            // If we have user mappings, try to find a mapping for this username
            if (userMappings != null && !string.IsNullOrWhiteSpace(userName))
            {
                var mapping = userMappings.FirstOrDefault(m =>
                    m.PrevUserName.Equals(userName, StringComparison.OrdinalIgnoreCase));

                if (mapping != null)
                {
                    return mapping.CurrentUserName ?? "*";
                }
            }

            return userName ?? "*";
        }

        return "*";
    }
    private static string ResolveUserOrGroup(JsonElement element, List<UserMapping> userMappings = null )
    {
        if (element.ValueKind != JsonValueKind.Object) return "";

        string resolvedName = "";

        if (element.TryGetProperty("Name", out var name) && name.ValueKind == JsonValueKind.String)
        {
            var val = name.GetString() ?? "";
            if (val.Contains("\\"))
            {
                resolvedName = val.Split('\\').Last();
            }
            else if (element.TryGetProperty("Title", out var title1) && title1.ValueKind == JsonValueKind.String)
            {
                resolvedName = title1.GetString().Split('\\').Last() ?? "";
            }
        }
        else if (element.TryGetProperty("Title", out var title) && title.ValueKind == JsonValueKind.String)
        {
            resolvedName = title.GetString() ?? "";
        }

        // Apply user mapping if available and we have a resolved name
        if (!string.IsNullOrEmpty(resolvedName) && userMappings != null)
        {
            var mapping = userMappings.FirstOrDefault(m =>
                m.PrevUserName.Equals(resolvedName, StringComparison.OrdinalIgnoreCase));

            if (mapping != null)
            {
                return mapping.CurrentUserName;
            }
        }

        return resolvedName;
    }
    public static string[] ResolveUsersOrGroups(this JsonElement element, string propertyName, List<UserMapping> userMappings = null)
    {
        if (!element.TryGetProperty(propertyName, out var prop))
            return Array.Empty<string>();

        return prop.ValueKind switch
        {
            JsonValueKind.Array => prop.EnumerateArray()
                                       .SelectMany(u => ExtractUsersOrGroups(u, userMappings)) // flatten List<string>
                                       .Where(x => !string.IsNullOrWhiteSpace(x))
                                       .ToArray(),

            JsonValueKind.Object => ExtractUsersOrGroups(prop, userMappings)
                                       .Where(x => !string.IsNullOrWhiteSpace(x))
                                       .ToArray(),

            JsonValueKind.String => HandleStringUserOrGroup(prop.GetString(), userMappings),

            _ => Array.Empty<string>()
        };
    }

    private static string[] HandleStringUserOrGroup(string? str, List<UserMapping> userMappings = null)
    {
        if (string.IsNullOrWhiteSpace(str)) return Array.Empty<string>();

        // Sometimes SP serializes as JSON string
        if (str.TrimStart().StartsWith("{") || str.TrimStart().StartsWith("["))
        {
            try
            {
                using var doc = JsonDocument.Parse(str);
                // recursive on root
                return doc.RootElement.ResolveUsersOrGroups("", userMappings);
            }
            catch
            {
                return ApplyUserMapping(new[] { str }, userMappings);
            }
        }

        return ApplyUserMapping(new[] { str }, userMappings);
    }

    public static List<string> ExtractUsersOrGroups(JsonElement userElement, List<UserMapping> userMappings = null)
    {
        var users = new List<string>();

        if (userElement.ValueKind == JsonValueKind.Object)
        {
            // If SharePoint returned { "results": [ ... ] }
            if (userElement.TryGetProperty("results", out var results) && results.ValueKind == JsonValueKind.Array)
            {
                foreach (var u in results.EnumerateArray())
                {
                    var login = ExtractSingleUserOrGroup(u, userMappings);
                    if (!string.IsNullOrWhiteSpace(login))
                        users.Add(login);
                }
                return users;
            }

            // Otherwise assume it's a single object
            var single = ExtractSingleUserOrGroup(userElement, userMappings);
            if (!string.IsNullOrWhiteSpace(single))
                users.Add(single);
        }

        return users;
    }

    private static string ExtractSingleUserOrGroup(JsonElement user, List<UserMapping> userMappings = null)
    {
        if (user.ValueKind != JsonValueKind.Object) return "";

        string resolvedName = "";

        // Prefer Name (claims login or DOMAIN\username)
        if (user.TryGetProperty("Name", out var name) && name.ValueKind == JsonValueKind.String)
        {
            var val = name.GetString() ?? "";
            if (!string.IsNullOrWhiteSpace(val))
            {
                if (val.Contains("\\"))
                {
                    resolvedName = val.Split('\\').Last();
                }
                else
                {
                    // Fallback: Title (display name → check if it contains "\")
                    if (user.TryGetProperty("Title", out var t) && t.ValueKind == JsonValueKind.String)
                    {
                        var val2 = t.GetString() ?? "";
                        resolvedName = val2.Contains("\\") ? val2.Split('\\').Last() : val2;
                    }
                }
            }
        }

        // Fallback: LoginName (sometimes exists)
        if (string.IsNullOrEmpty(resolvedName) &&
            user.TryGetProperty("LoginName", out var login) && login.ValueKind == JsonValueKind.String)
        {
            var val = login.GetString() ?? "";
            resolvedName = val.Contains("\\") ? val.Split('\\').Last() : val;
        }

        // Fallback: Title (if not already used)
        if (string.IsNullOrEmpty(resolvedName) &&
            user.TryGetProperty("Title", out var title) && title.ValueKind == JsonValueKind.String)
        {
            resolvedName = title.GetString() ?? "";
        }

        // Apply user mapping if available
        return ApplyUserMapping(resolvedName, userMappings);
    }

    private static string ExtractUserOrGroup(JsonElement userElement, List<UserMapping> userMappings = null)
    {
        if (userElement.ValueKind != JsonValueKind.Object) return "";

        string resolvedName = "";

        if (userElement.TryGetProperty("Name", out var name) && name.ValueKind == JsonValueKind.String)
        {
            var val = name.GetString() ?? "";
            resolvedName = val.Contains("\\") ? val.Split('\\').Last() : val;
        }

        if (string.IsNullOrEmpty(resolvedName) &&
            userElement.TryGetProperty("LoginName", out var login) && login.ValueKind == JsonValueKind.String)
        {
            resolvedName = login.GetString() ?? "";
        }

        if (string.IsNullOrEmpty(resolvedName) &&
            userElement.TryGetProperty("Title", out var title) && title.ValueKind == JsonValueKind.String)
        {
            resolvedName = title.GetString() ?? "";
        }

        // Apply user mapping if available
        return ApplyUserMapping(resolvedName, userMappings);
    }

    // Helper method to apply user mapping to a single username
    private static string ApplyUserMapping(string username, List<UserMapping> userMappings = null)
    {
        if (string.IsNullOrWhiteSpace(username)) return username;

        // Apply user mapping if available
        if (userMappings != null)
        {
            var mapping = userMappings.FirstOrDefault(m =>
                m.PrevUserName.Equals(username, StringComparison.OrdinalIgnoreCase));

            if (mapping != null)
            {
                return mapping.CurrentUserName;
            }
        }

        return username;
    }

    // Helper method to apply user mapping to an array of usernames
    private static string[] ApplyUserMapping(string[] usernames, List<UserMapping> userMappings = null)
    {
        if (usernames == null || usernames.Length == 0) return usernames;
        if (userMappings == null) return usernames;

        var result = new string[usernames.Length];
        for (int i = 0; i < usernames.Length; i++)
        {
            result[i] = ApplyUserMapping(usernames[i], userMappings);
        }
        return result;
    }
}

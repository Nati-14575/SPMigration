using Microsoft.AspNetCore.StaticFiles;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.RecordsRepository;
using Minio;
using Minio.ApiEndpoints;
using Minio.DataModel.Args;
using SPMigration.Models;

public sealed class MinIOService
{
    private readonly IMinioClient _minio;
    private readonly string _bucket;
    private readonly FileExtensionContentTypeProvider _mime = new();
    private readonly string _unitKey;
    
    public MinIOService(IConfiguration config)
    {
        var endpoint = config["MinIO:Endpoint"];
        var accessKey = config["MinIO:AccessKey"];
        var secretKey = config["MinIO:SecretKey"];
        var useSsl = config["MinIO:UseSSL"];

        _bucket = config["MinIO:Bucket"];
        _unitKey = config["MinIO:UnitKey"];

        _minio = new Minio.MinioClient()
            .WithEndpoint(endpoint)
            .WithSSL(bool.TryParse(useSsl, out var sslEnabled) && sslEnabled)
            .WithCredentials(accessKey, secretKey)
            .Build();
    }

    public async Task EnsureBucketAsync(CancellationToken ct = default)
    {
        var exists = await _minio.BucketExistsAsync(new BucketExistsArgs().WithBucket(_bucket), ct);
        if (!exists)
            await _minio.MakeBucketAsync(new MakeBucketArgs().WithBucket(_bucket), ct);
    }

    public sealed record FileObj(string bucket, string fileName, string contentType, long size, string originalname);

    public async Task<(MinIOFileInfo fileInfo, List<MinIOFileInfo> attachments)>
     UploadListItemAttachmentsAsync(int Id, List<(string FileName, byte[] Bytes, string contentType)> files, string recordTitle)
    {
        MinIOFileInfo main = null;
        var atts = new List<MinIOFileInfo>();

        int index = 0;
        foreach (var (fileName, bytes, contentType) in files)
        {
            // Check normal rule
            var isRecord = fileName.Equals($"{recordTitle}.pdf", StringComparison.OrdinalIgnoreCase);

            // Fallback: if main not set yet and this is the first file, mark as record
            if (!isRecord && main == null && index == 0)
            {
                isRecord = true;
            }
            var fileTitle = SanitizeObjectName(fileName);
            var key = isRecord
            ? $"migrated/{_unitKey}/record-{Id}/records/{fileTitle}"
            : $"migrated/{_unitKey}/record-{Id}/attachments/{fileTitle}";


            using var mem = new MemoryStream(bytes);
            await _minio.PutObjectAsync(new PutObjectArgs()
                .WithBucket(_bucket)
                .WithObject(key)
                .WithStreamData(mem)
                .WithObjectSize(mem.Length)
                .WithContentType(contentType)); // TODO: detect MIME type if needed

            var obj = new MinIOFileInfo
            {
                Bucket = _bucket,
                FileName = key,
                ContentType = contentType,
                Size = mem.Length,
                Originalname = fileName
            };

            if (isRecord) main = obj;
            else atts.Add(obj);

            index++;
        }

        // If no files at all, main stays null
        return (main, atts);

    }
    
    public async Task<(MinIOFileInfo fileInfo, List<MinIOFileInfo> attachments)>
     UploadFilesAsync(int Id, List<(string FileName, byte[] Bytes, string contentType)> files, string recordTitle)
    {
        MinIOFileInfo main = null;
        var atts = new List<MinIOFileInfo>();

        int index = 0;
        foreach (var (fileName, bytes, contentType) in files)
        {
            // Check normal rule
            var isRecord = fileName.Equals($"{recordTitle}.pdf", StringComparison.OrdinalIgnoreCase);

            // Fallback: if main not set yet and this is the first file, mark as record
            if (!isRecord && main == null && index == 0)
            {
                isRecord = true;
            }
            var fileTitle = SanitizeObjectName(fileName);
            var key = isRecord
            ? $"migrated/{_unitKey}/record-{Id}/records/{fileTitle}"
            : $"migrated/{_unitKey}/record-{Id}/attachments/{fileTitle}";


            using var mem = new MemoryStream(bytes);
            await _minio.PutObjectAsync(new PutObjectArgs()
                .WithBucket(_bucket)
                .WithObject(key)
                .WithStreamData(mem)
                .WithObjectSize(mem.Length)
                .WithContentType(contentType)); // TODO: detect MIME type if needed

            var obj = new MinIOFileInfo
            {
                Bucket = _bucket,
                FileName = key,
                ContentType = contentType,
                Size = mem.Length,
                Originalname = fileName
            };

            if (isRecord) main = obj;
            else atts.Add(obj);

            index++;
        }

        // If no files at all, main stays null
        return (main, atts);

    }

    public async Task<(MinIOFileInfo fileInfo, List<MinIOFileInfo> attachments)>
    GetRecordFilesAsync(int recordId, CancellationToken ct = default)
    {
        var prefix = $"migrated/{_unitKey}/record-{recordId}/";
        var files = new List<MinIOFileInfo>();

        var observable = _minio.ListObjectsAsync(
            new ListObjectsArgs()
                .WithBucket(_bucket)
                .WithPrefix(prefix)
                .WithRecursive(true),
            ct);

        var tcs = new TaskCompletionSource<bool>();
        observable.Subscribe(
            item =>
            {
                files.Add(new MinIOFileInfo
                {
                    Bucket = _bucket,
                    FileName = item.Key,
                    ContentType = item.ContentType ?? "application/octet-stream",
                    Size = (long)item.Size,
                    Originalname = Path.GetFileName(item.Key)
                });
            },
            ex => tcs.TrySetException(ex),
            () => tcs.TrySetResult(true));

        await tcs.Task;

        if (files.Count == 0)
            return (null, new List<MinIOFileInfo>());

        var main = files.FirstOrDefault(f => f.FileName.Contains("/records/")) ?? files.First();
        var atts = files.Where(f => f != main).ToList();

        return (main, atts);
    }
    public async Task<(MinIOFileInfo fileInfo, List<MinIOFileInfo> attachments)>
    GetRecordFilesFromFolderAsync(int recordId, CancellationToken ct = default)
    {
        var prefix = $"migrated/externalIncoming/{_unitKey}/record-{recordId}/";
        var files = new List<MinIOFileInfo>();

        var observable = _minio.ListObjectsAsync(
            new ListObjectsArgs()
                .WithBucket(_bucket)
                .WithPrefix(prefix)
                .WithRecursive(true),
            ct);

        var tcs = new TaskCompletionSource<bool>();
        observable.Subscribe(
            item =>
            {
                files.Add(new MinIOFileInfo
                {
                    Bucket = _bucket,
                    FileName = item.Key,
                    ContentType = item.ContentType ?? "application/octet-stream",
                    Size = (long)item.Size,
                    Originalname = Path.GetFileName(item.Key)
                });
            },
            ex => tcs.TrySetException(ex),
            () => tcs.TrySetResult(true));

        await tcs.Task;

        if (files.Count == 0)
            return (null, new List<MinIOFileInfo>());

        var main = files.FirstOrDefault(f => f.FileName.Contains("/records/")) ?? files.First();
        var atts = files.Where(f => f != main).ToList();

        return (main, atts);
    }
    private string SanitizeObjectName(string input)
    {
        // Replace unsupported chars with "-"
        var invalid = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
        foreach (var c in invalid)
        {
            input = input.Replace(c, '-');
        }

        // Optional: remove spaces and trim
        input = input.Replace(" ", "_").Trim();
        input = input.Replace("•", "_").Trim();

        return input;
    }


}

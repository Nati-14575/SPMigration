public static class MigrationErrorTypes
{
    /// <summary>The record/folder payload itself was rejected or processing failed.</summary>
    public const string Migration = "Migration";
    /// <summary>The record migrated, but updating its file info/attachments failed.</summary>
    public const string FileInfo = "FileInfo";
}

public class MigrationException : Exception
{
    public object Payload { get; }
    public string ErrorType { get; }

    public MigrationException(string message, object payload, Exception innerException = null, string errorType = MigrationErrorTypes.Migration)
        : base(message, innerException)
    {
        Payload = payload;
        ErrorType = errorType;
    }
}

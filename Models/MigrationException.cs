public class MigrationException : Exception
{
    public object Payload { get; }

    public MigrationException(string message, object payload, Exception innerException = null)
        : base(message, innerException)
    {
        Payload = payload;
    }
}

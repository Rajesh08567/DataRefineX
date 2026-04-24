namespace DataRefineX.Models;

public enum LogLevel
{
    Info,
    Success,
    Warning,
    Error
}

public sealed class LogEntry
{
    public LogEntry(LogLevel level, string message)
    {
        Level = level;
        Message = message;
        Timestamp = DateTime.Now;
    }

    public DateTime Timestamp { get; }
    public LogLevel Level { get; }
    public string Message { get; }

    public string TimeDisplay => Timestamp.ToString("HH:mm:ss");
}

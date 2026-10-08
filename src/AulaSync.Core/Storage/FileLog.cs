namespace AulaSync.Core;

public sealed class FileLog(string path, long maxBytes = 1_000_000)
{
    readonly object _lock = new();

    public void Info(string message) => Write("INFO", message);

    public void Error(string message, Exception? ex = null) => Write("FEJL", ex is null ? message : $"{message}: {ex}");

    void Write(string level, string message)
    {
        lock (_lock)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
                if (File.Exists(path) && new FileInfo(path).Length > maxBytes)
                    File.Move(path, path + ".old", overwrite: true);
                File.AppendAllText(path, $"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss} {level} {message}{Environment.NewLine}");
            }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }
}

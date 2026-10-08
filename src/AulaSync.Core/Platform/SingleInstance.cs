namespace AulaSync.Core;

// Eksklusiv fillås. Virker på Windows og Mac og slippes automatisk, hvis processen dør.
public sealed class SingleInstance : IDisposable
{
    readonly FileStream _lock;

    SingleInstance(FileStream stream) => _lock = stream;

    public static SingleInstance? TryAcquire(string lockPath)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(lockPath))!);
        try { return new SingleInstance(new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None)); }
        catch (IOException) { return null; }
    }

    public void Dispose() => _lock.Dispose();
}

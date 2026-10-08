using System.Text;

namespace AulaSync.Core;

public static class AtomicFile
{
    static readonly UTF8Encoding Utf8NoBom = new(encoderShouldEmitUTF8Identifier: false);

    // Skriv til midlertidig fil i samme mappe og flyt over den gamle.
    // Windows kan afvise flytningen et kort øjeblik, hvis en anden læser filen — derfor nogle få forsøg.
    public static void WriteAllText(string path, string content)
    {
        var full = Path.GetFullPath(path);
        Directory.CreateDirectory(Path.GetDirectoryName(full)!);
        var tmp = $"{full}.{Guid.NewGuid():N}.tmp";
        try
        {
            using (var stream = new FileStream(tmp, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                stream.Write(Utf8NoBom.GetBytes(content));
                stream.Flush(flushToDisk: true);
            }
            for (int attempt = 1; ; attempt++)
            {
                try { File.Move(tmp, full, overwrite: true); return; }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException && attempt < 10)
                {
                    Thread.Sleep(100);
                }
            }
        }
        finally
        {
            if (File.Exists(tmp)) File.Delete(tmp);
        }
    }
}

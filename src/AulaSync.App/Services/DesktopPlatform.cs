using System.Diagnostics;
using System.Runtime.InteropServices;
using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Input.Platform;
using AulaSync.Core;

namespace AulaSync.App;

// setClipboard: kun til testene; ellers Avalonias udklipsholder.
public sealed class DesktopPlatform(FileLog log, Func<string, Task>? setClipboard = null) : IPlatform
{
    public bool IsMac => OperatingSystem.IsMacOS();

    public void Open(string target) => Start(new ProcessStartInfo(target) { UseShellExecute = true });

    // "open -b com.apple.iCal" sender adressen til Kalender (ikke til browseren, som ellers åbner http).
    public void OpenInCalendar(string url)
    {
        if (IsMac) Start(new ProcessStartInfo("open", ["-b", "com.apple.iCal", url]));
        else Open(url);
    }

    public void RevealFile(string path)
    {
        if (IsMac) Start(new ProcessStartInfo("open", ["-R", path]));
        else if (OperatingSystem.IsWindows()) Start(new ProcessStartInfo("explorer.exe", $"/select,\"{path}\""));
        else Open(Path.GetDirectoryName(path)!);
    }

    // Windows: holder et andet program udklipsholderen i mere end ca. et sekund (fjernskrivebord, Office, en
    // udklipsholder-app), kaster Avalonia en COMException. Den må ikke lukke appen.
    public async Task<bool> CopyTextAsync(string text)
    {
        try
        {
            await (setClipboard ?? SetClipboardAsync)(text);
            return true;
        }
        catch (Exception ex) when (ex is ExternalException or TimeoutException)
        {
            log.Error("Kunne ikke kopiere til udklipsholderen", ex);
            return false;
        }
    }

    static async Task SetClipboardAsync(string text)
    {
        var lifetime = Application.Current?.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime;
        var window = lifetime?.Windows.FirstOrDefault(w => w.IsActive) ?? lifetime?.Windows.FirstOrDefault();
        if (window?.Clipboard is { } clipboard) await clipboard.SetTextAsync(text);
    }

    void Start(ProcessStartInfo info)
    {
        try { using var _ = Process.Start(info); }
        catch (Exception ex) { log.Error($"Kunne ikke starte {info.FileName}", ex); }
    }
}

using System.Runtime.Versioning;
using AulaSync.Core;
using Microsoft.Win32;

namespace AulaSync.App;

// "Start AulaSync, når jeg logger ind" (spec §3.3). Starter med --silent: intet vindue, kun ikonet.
public interface IAutostart
{
    bool IsEnabled { get; }
    void SetEnabled(bool enabled);

    // Er start ved login slået til, men peger på en anden AulaSync end den, der kører (fx den løse AulaSync.exe fra før
    // installationsprogrammet), skrives startpunktet igen. true, hvis det blev skrevet.
    bool Retarget();
}

public sealed class MacAutostart(string home, string executablePath) : IAutostart
{
    string PlistPath => LaunchAgent.PlistPath(home);

    public bool IsEnabled => File.Exists(PlistPath);

    public void SetEnabled(bool enabled)
    {
        if (enabled) AtomicFile.WriteAllText(PlistPath, LaunchAgent.CreatePlist(executablePath));
        else File.Delete(PlistPath);
    }

    // Kører appen fra .dmg-filen (/Volumes) eller fra macOS' midlertidige kopi (AppTranslocation), forsvinder stien, når
    // .dmg'en skubbes ud; så røres startpunktet ikke.
    public bool Retarget()
    {
        if (IsTransient(executablePath) || !IsEnabled || File.ReadAllText(PlistPath) == LaunchAgent.CreatePlist(executablePath)) return false;
        SetEnabled(true);
        return true;
    }

    internal static bool IsTransient(string path) =>
        path.StartsWith("/Volumes/", StringComparison.Ordinal) || path.Contains("/AppTranslocation/", StringComparison.Ordinal);
}

[SupportedOSPlatform("windows")]
public sealed class WindowsAutostart(string executablePath, string runKey = WindowsAutostart.RunKey) : IAutostart
{
    public const string RunKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
    const string ValueName = "AulaSync";

    string Command => $"\"{executablePath}\" --silent";

    string? Current
    {
        get
        {
            using var key = Registry.CurrentUser.OpenSubKey(runKey);
            return key?.GetValue(ValueName) as string;
        }
    }

    public bool IsEnabled => Current is not null;

    public void SetEnabled(bool enabled)
    {
        using var key = Registry.CurrentUser.CreateSubKey(runKey);
        if (enabled) key.SetValue(ValueName, Command);
        else key.DeleteValue(ValueName, throwOnMissingValue: false);
    }

    public bool Retarget()
    {
        if (Current is not { } current || string.Equals(current, Command, StringComparison.OrdinalIgnoreCase)) return false;
        SetEnabled(true);
        return true;
    }
}

public sealed class NoAutostart : IAutostart
{
    public bool IsEnabled => false;
    public void SetEnabled(bool enabled) { }
    public bool Retarget() => false;
}

public static class AutostartFactory
{
    public static IAutostart ForCurrentPlatform(string executablePath) =>
        OperatingSystem.IsWindows() ? new WindowsAutostart(executablePath)
        : OperatingSystem.IsMacOS() ? new MacAutostart(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), executablePath)
        : new NoAutostart();
}

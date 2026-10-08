using System.Runtime.Versioning;
using AulaSync.Core;
using Microsoft.Win32;

namespace AulaSync.App;

// "Start AulaSync, når jeg logger ind" (spec §3.3). Starter med --silent: intet vindue, kun ikonet.
public interface IAutostart
{
    bool IsEnabled { get; }
    void SetEnabled(bool enabled);
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
}

[SupportedOSPlatform("windows")]
public sealed class WindowsAutostart(string executablePath) : IAutostart
{
    const string RunKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
    const string ValueName = "AulaSync";

    public bool IsEnabled
    {
        get
        {
            using var key = Registry.CurrentUser.OpenSubKey(RunKey);
            return key?.GetValue(ValueName) is string;
        }
    }

    public void SetEnabled(bool enabled)
    {
        using var key = Registry.CurrentUser.CreateSubKey(RunKey);
        if (enabled) key.SetValue(ValueName, $"\"{executablePath}\" --silent");
        else key.DeleteValue(ValueName, throwOnMissingValue: false);
    }
}

public sealed class NoAutostart : IAutostart
{
    public bool IsEnabled => false;
    public void SetEnabled(bool enabled) { }
}

public static class AutostartFactory
{
    public static IAutostart ForCurrentPlatform(string executablePath) =>
        OperatingSystem.IsWindows() ? new WindowsAutostart(executablePath)
        : OperatingSystem.IsMacOS() ? new MacAutostart(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), executablePath)
        : new NoAutostart();
}

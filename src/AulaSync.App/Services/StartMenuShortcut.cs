using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.Versioning;
using System.Text;

namespace AulaSync.App;

// Windows: genvejen "AulaSync" i brugerens Start-menu, så AulaSync kan findes i Start og i søgningen. winget laver ingen
// genvej til en portabel app, og AulaSync 2 lavede en med samme navn, der peger på den gamle exe. Ved hver start skrives
// genvejen, hvis den mangler eller peger på en anden fil end den, der kører; ellers røres den ikke (en fastgørelse i Start
// følger filen, så den bliver ved med at virke).
public static class StartMenuShortcut
{
    public const string FileName = "AulaSync.lnk";
    public const string Description = "Holder skemaer fra Aula opdateret i din kalender";

    // %APPDATA%\Microsoft\Windows\Start Menu\Programs: kun brugerens egen Start-menu, så det kræver ikke administrator.
    public static string DefaultPath => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Programs), FileName);

    // Windows skelner ikke mellem store og små bogstaver i stier.
    public static bool PointsTo(string? target, string exePath) =>
        !string.IsNullOrEmpty(target) && string.Equals(target, exePath, StringComparison.OrdinalIgnoreCase);

    // Startes AulaSync gennem winget's henvisning (Links\AulaSync.exe), peger genvejen på selve filen, så den er den samme,
    // uanset hvordan AulaSync blev startet. Kun ét led: det endelige mål kan på Windows komme tilbage som en forkert sti for
    // filer på et netværksdrev. Findes målet ikke, bruges stien, AulaSync blev startet med.
    public static string RealPath(string exePath)
    {
        try
        {
            var real = new FileInfo(exePath).ResolveLinkTarget(returnFinalTarget: false)?.FullName;
            return real is not null && File.Exists(real) ? real : exePath;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { return exePath; }
    }

    // true, hvis genvejen blev skrevet.
    [SupportedOSPlatform("windows")]
    public static bool Ensure(string linkPath, string exePath)
    {
        if (File.Exists(linkPath) && PointsTo(TryReadTarget(linkPath), exePath)) return false;
        Directory.CreateDirectory(Path.GetDirectoryName(linkPath)!);
        var link = (IShellLinkW)new ShellLink();
        try
        {
            link.SetPath(exePath);
            link.SetWorkingDirectory(Path.GetDirectoryName(exePath)!);
            link.SetDescription(Description);
            ((IPersistFile)link).Save(linkPath, true);
        }
        finally { Marshal.FinalReleaseComObject(link); }
        return true;
    }

    // En ødelagt genvej kan ikke læses; så skrives den forfra.
    [SupportedOSPlatform("windows")]
    static string? TryReadTarget(string linkPath)
    {
        try { return ReadTarget(linkPath); }
        catch (Exception ex) when (ex is COMException or IOException or UnauthorizedAccessException) { return null; }
    }

    // Filen, genvejen peger på, eller null, hvis den ikke peger på en fil.
    [SupportedOSPlatform("windows")]
    public static string? ReadTarget(string linkPath)
    {
        var link = (IShellLinkW)new ShellLink();
        try
        {
            ((IPersistFile)link).Load(linkPath, 0);
            var path = new StringBuilder(32767);
            link.GetPath(path, path.Capacity, 0, 0);
            return path.Length == 0 ? null : path.ToString();
        }
        finally { Marshal.FinalReleaseComObject(link); }
    }

    [ComImport, Guid("00021401-0000-0000-C000-000000000046")]
    class ShellLink { }

    // Rækkefølgen skal følge IShellLinkW i shobjidl_core.h.
    [ComImport, Guid("000214F9-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IShellLinkW
    {
        void GetPath([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFile, int cch, nint pfd, uint fFlags);
        void GetIDList(out nint ppidl);
        void SetIDList(nint pidl);
        void GetDescription([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszName, int cch);
        void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string pszName);
        void GetWorkingDirectory([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszDir, int cch);
        void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string pszDir);
        void GetArguments([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszArgs, int cch);
        void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string pszArgs);
        void GetHotkey(out ushort pwHotkey);
        void SetHotkey(ushort wHotkey);
        void GetShowCmd(out int piShowCmd);
        void SetShowCmd(int iShowCmd);
        void GetIconLocation([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszIconPath, int cch, out int piIcon);
        void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string pszIconPath, int iIcon);
        void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string pszPathRel, uint dwReserved);
        void Resolve(nint hwnd, uint fFlags);
        void SetPath([MarshalAs(UnmanagedType.LPWStr)] string pszFile);
    }
}

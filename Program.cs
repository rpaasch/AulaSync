using System.Runtime.InteropServices;

namespace AulaSync;

static class Program
{
    static NotifyIcon Tray = null!;
    public static Form HiddenForm { get; private set; } = null!;

    public static bool Silent { get; set; }

    [STAThread]
    static void Main(string[] args)
    {
        // Kun én instans ad gangen
        using var mutex = new Mutex(true, "AulaSync_SingleInstance", out var isNew);
        if (!isNew) return;

        Silent = args.Any(a => a is "--silent" or "/silent" or "-silent");
        EnsureStartMenuShortcut();
        Application.SetHighDpiMode(HighDpiMode.PerMonitorV2);
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        var icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath) ?? SystemIcons.Application;
        TrayApp.AppIcon = icon;

        // Opret tray-ikon FØRST (synkront, på UI-tråden)
        Tray = new NotifyIcon
        {
            Icon = icon,
            Text = "AulaSync - starter...",
            Visible = true,
        };

        // Skjult form som message pump ejer (giver os et window handle)
        HiddenForm = new Form
        {
            FormBorderStyle = FormBorderStyle.None,
            ShowInTaskbar = false,
            Opacity = 0,
            Size = new Size(0, 0),
        };

        HiddenForm.Shown += async (_, _) =>
        {
            HiddenForm.Hide();
            var app = new TrayApp(Tray);
            await app.RunLoopAsync();
        };

        Application.Run(HiddenForm);
    }

    /// <summary>
    /// Opretter startmenu-genvej i brugerens egen startmenu (kræver ikke admin).
    /// </summary>
    private static void EnsureStartMenuShortcut()
    {
        try
        {
            var startMenu = Environment.GetFolderPath(Environment.SpecialFolder.StartMenu);
            var lnkPath = Path.Combine(startMenu, "Programs", "AulaSync.lnk");
            if (File.Exists(lnkPath)) return;

            var shell = (IShellLinkW)new ShellLink();
            shell.SetPath(Application.ExecutablePath);
            // Ingen --silent: brugeren skal se login ved manuelt klik
            shell.SetDescription("Synkroniserer Aula til Outlook");
            shell.SetWorkingDirectory(Path.GetDirectoryName(Application.ExecutablePath)!);

            var file = (IPersistFile)shell;
            file.Save(lnkPath, false);
        }
        catch { }
    }

    [ComImport, Guid("00021401-0000-0000-C000-000000000046")]
    private class ShellLink { }

    [ComImport, Guid("000214F9-0000-0000-C000-000000000046"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IShellLinkW
    {
        void GetPath(nint pszFile, int cch, nint pfd, uint fFlags);
        void GetIDList(out nint ppidl);
        void SetIDList(nint pidl);
        void GetDescription(nint pszName, int cch);
        void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string pszName);
        void GetWorkingDirectory(nint pszDir, int cch);
        void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string pszDir);
        void GetArguments(nint pszArgs, int cch);
        void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string pszArgs);
        void GetHotkey(out ushort pwHotkey);
        void SetHotkey(ushort wHotkey);
        void GetShowCmd(out int piShowCmd);
        void SetShowCmd(int iShowCmd);
        void GetIconLocation(nint pszIconPath, int cch, out int piIcon);
        void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string pszIconPath, int iIcon);
        void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string pszPathRel, uint dwReserved);
        void Resolve(nint hwnd, uint fFlags);
        void SetPath([MarshalAs(UnmanagedType.LPWStr)] string pszFile);
    }

    [ComImport, Guid("0000010b-0000-0000-C000-000000000046"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IPersistFile
    {
        void GetClassID(out Guid pClassID);
        int IsDirty();
        void Load([MarshalAs(UnmanagedType.LPWStr)] string pszFileName, uint dwMode);
        void Save([MarshalAs(UnmanagedType.LPWStr)] string pszFileName, [MarshalAs(UnmanagedType.Bool)] bool fRemember);
        void SaveCompleted([MarshalAs(UnmanagedType.LPWStr)] string pszFileName);
        void GetCurFile(out nint ppszFileName);
    }
}

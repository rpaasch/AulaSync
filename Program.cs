namespace AulaSync;

static class Program
{
    static NotifyIcon Tray = null!;
    public static Form HiddenForm { get; private set; } = null!;

    public static bool Silent { get; set; }

    [STAThread]
    static void Main(string[] args)
    {
        Silent = args.Any(a => a is "--silent" or "/silent" or "-silent");
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
}

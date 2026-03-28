using System.Text.Json;
using Microsoft.Win32;

namespace AulaSync;

class TrayApp
{
    public static Icon AppIcon { get; set; } = null!;

    private readonly NotifyIcon _tray;
    private AulaApi? _api;
    private HttpClient? _httpClient;
    private System.Windows.Forms.Timer _pollTimer;
    private int _totalImported;
    private DateTime? _lastPoll;
    private string? _lastError;
    private string _header = "AulaSync";

    private static readonly string ConfigDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".aulasync");
    private static readonly string ConfigFile = Path.Combine(ConfigDir, "config.json");
    private static readonly string LogFile = Path.Combine(ConfigDir, "aulasync.log");

    private static readonly Dictionary<string, int> PollIntervals = new()
    {
        ["Hvert 2. minut"] = 120_000,
        ["Hvert 5. minut"] = 300_000,
        ["Hvert 10. minut"] = 600_000,
        ["Hvert 30. minut"] = 1_800_000,
        ["Hver time"] = 3_600_000,
    };

    private const string AutostartKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
    private const string AutostartName = "AulaSync";
    private const string ProjectUrl = "https://github.com/rpaasch/AulaSync";

    private TaskCompletionSource? _currentWait;
    private bool _quitRequested;

    public TrayApp(NotifyIcon tray)
    {
        Directory.CreateDirectory(ConfigDir);
        _tray = tray;
        _tray.MouseClick += (_, e) =>
        {
            Log($"Tray klik: {e.Button}");
            if (_tray.ContextMenuStrip != null)
            {
                Log($"Menu items: {_tray.ContextMenuStrip.Items.Count}");
                // Force vis menuen
                var mi = typeof(NotifyIcon).GetMethod("ShowContextMenu",
                    System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
                mi?.Invoke(_tray, null);
                Log("ShowContextMenu kaldt");
            }
            else
            {
                Log("ContextMenuStrip er NULL!");
            }
        };

        _pollTimer = new System.Windows.Forms.Timer();
        _pollTimer.Interval = LoadConfig().TryGetValue("poll_interval", out var pi) ? (int)pi : 300_000;
        _pollTimer.Tick += async (_, _) => await DoFetchAsync();
    }

    public async Task RunLoopAsync()
    {
        while (true)
        {
            _tray.Text = "AulaSync - logger ind...";

            Log("Starter login...");
            var login = new AulaLogin();
            if (!await login.LoginAsync())
            {
                _tray.Visible = false;
                _tray.Dispose();
                Application.Exit();
                return;
            }

            Log("Login OK - opretter API-klient...");
            _httpClient = login.CreateHttpClient();
            _api = new AulaApi(_httpClient);

            if (!await _api.ConnectAsync())
            {
                Log("ConnectAsync fejlede - prøver igen");
                continue;
            }

            _header = string.Join(" - ", new[] { _api.UserName, _api.RoleDanish, _api.Institution }
                .Where(s => !string.IsNullOrEmpty(s)));
            _tray.Text = $"AulaSync - {_api.UserName}";

            // Sæt menuen med det samme
            _tray.ContextMenuStrip = BuildMenu();
            ShowBalloon($"Logget ind som {_api.UserName}");
            Log($"Logget ind som {_api.UserName}");

            // Hent beskeder i baggrunden (blokerer ikke UI)
            _ = DoFetchAsync();

            // Start polling
            _pollTimer.Start();

            // Vent på logout/quit
            _currentWait = new TaskCompletionSource();
            await _currentWait.Task;
            _currentWait = null;

            _pollTimer.Stop();

            if (_quitRequested)
            {
                _tray.Visible = false;
                _tray.Dispose();
                Application.Exit();
                return;
            }

            // Logout → loop
            _api = null;
            _httpClient?.Dispose();
            _httpClient = null;
        }
    }

    private ContextMenuStrip BuildMenu()
    {
        var menu = new ContextMenuStrip();

        int eggClicks = 0;
        var h = menu.Items.Add(_header);
        h.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
        var credit = new ToolStripMenuItem("af rpaasch") { Visible = false };
        credit.Click += (_, _) =>
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            { FileName = ProjectUrl, UseShellExecute = true });
        };
        h.Click += (_, _) =>
        {
            eggClicks++;
            if (eggClicks >= 5) credit.Visible = true;
        };
        menu.Closing += (_, e) =>
        {
            if (eggClicks > 0 && eggClicks < 5 && e.CloseReason == ToolStripDropDownCloseReason.ItemClicked)
            { e.Cancel = true; }
        };

        menu.Items.Add(credit);
        var s = menu.Items.Add(StatusText());
        s.Enabled = false;
        menu.Opening += (_, _) => s.Text = StatusText();

        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add("Hent nu", null, async (_, _) => await DoFetchAsync());

        menu.Items.Add(new ToolStripSeparator());

        // Indstillinger submenu
        var settings = new ToolStripMenuItem("Indstillinger");

        var interval = new ToolStripMenuItem("Synkroniseringsinterval");
        foreach (var (label, ms) in PollIntervals)
        {
            var c = ms;
            var item = new ToolStripMenuItem(label, null, (_, _) =>
            {
                _pollTimer.Interval = c;
                SaveConfigValue("poll_interval", c);
            });
            interval.DropDownOpening += (_, _) => item.Checked = _pollTimer.Interval == c;
            interval.DropDownItems.Add(item);
        }
        settings.DropDownItems.Add(interval);

        var auto = new ToolStripMenuItem("Start ved login", null, (_, _) =>
            SetAutostart(!IsAutostartEnabled()));
        menu.Opening += (_, _) => auto.Checked = IsAutostartEnabled();
        settings.DropDownItems.Add(auto);

        settings.DropDownItems.Add(new ToolStripSeparator());
        settings.DropDownItems.Add("Nulstil importhistorik", null, (_, _) =>
        {
            OutlookBridge.ResetSeen();
            ShowBalloon("Importhistorik nulstillet");
        });
        settings.DropDownItems.Add("Vis log", null, (_, _) =>
        {
            if (File.Exists(LogFile))
                System.Diagnostics.Process.Start("notepad.exe", LogFile);
        });
        settings.DropDownItems.Add("Vejledning", null, (_, _) =>
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            { FileName = $"{ProjectUrl}#vejledning", UseShellExecute = true });
        });
        var about = new ToolStripMenuItem("AulaSync v2.0.0") { Enabled = false };
        settings.DropDownItems.Add(about);

        menu.Items.Add(settings);

        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add("Log ud", null, (_, _) =>
        {
            _quitRequested = false;
            try { Directory.Delete(Path.Combine(ConfigDir, "webview2"), true); } catch { }
            _currentWait?.TrySetResult();
        });
        menu.Items.Add("Afslut", null, (_, _) =>
        {
            _quitRequested = true;
            _currentWait?.TrySetResult();
        });

        return menu;
    }

    private string StatusText()
    {
        var parts = new List<string> { $"{_totalImported} importeret" };
        if (_lastPoll.HasValue) parts.Add($"sidst {_lastPoll.Value:HH:mm}");
        if (_lastError != null) parts.Add($"fejl: {_lastError}");
        return string.Join(" · ", parts);
    }

    private async Task DoFetchAsync()
    {
        try
        {
            await _api!.TestSessionAsync();
            var count = await FetchMessagesAsync();
            _totalImported += count;
            _lastPoll = DateTime.Now;
            _lastError = null;
            if (count > 0) ShowBalloon($"{count} nye Aula-beskeder i Outlook");
        }
        catch (UnauthorizedAccessException)
        {
            _lastError = "Session udløbet";
            ShowBalloon("Session udløbet - log ind igen");
            _quitRequested = false;
            _currentWait?.TrySetResult();
        }
        catch (Exception ex)
        {
            _lastError = ex.Message;
            Log($"Fejl: {ex}");
        }
    }

    private async Task<int> FetchMessagesAsync()
    {
        var threads = await _api!.GetRecentThreadsAsync();
        var messages = new List<AulaMessage>();
        foreach (var t in threads)
        {
            var msg = await _api.GetFullMessageAsync(t);
            if (msg != null) messages.Add(msg);
        }
        // Outlook COM kører på baggrundstråd (STA) så UI ikke blokeres
        return await Task.Run(() => OutlookBridge.CreateOutlookItems(messages));
    }

    private void ShowBalloon(string text)
    {
        _tray.ShowBalloonTip(3000, "AulaSync", text, ToolTipIcon.Info);
    }

    private static void Log(string message)
    {
        try
        {
            Directory.CreateDirectory(ConfigDir);
            File.AppendAllText(LogFile, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {message}\n");
        }
        catch { }
    }

    private static Dictionary<string, long> LoadConfig()
    {
        try
        {
            if (File.Exists(ConfigFile))
                return JsonSerializer.Deserialize<Dictionary<string, long>>(
                    File.ReadAllText(ConfigFile)) ?? new();
        }
        catch { }
        return new();
    }

    private static void SaveConfigValue(string key, long value)
    {
        var config = LoadConfig();
        config[key] = value;
        Directory.CreateDirectory(ConfigDir);
        File.WriteAllText(ConfigFile,
            JsonSerializer.Serialize(config, new JsonSerializerOptions { WriteIndented = true }));
    }

    private static bool IsAutostartEnabled()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(AutostartKey, false);
            return key?.GetValue(AutostartName) != null;
        }
        catch { return false; }
    }

    private static void SetAutostart(bool enabled)
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(AutostartKey, true);
            if (key == null) return;
            if (enabled)
                key.SetValue(AutostartName, Application.ExecutablePath);
            else
                key.DeleteValue(AutostartName, false);
        }
        catch { }
    }
}

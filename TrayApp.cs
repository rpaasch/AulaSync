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
    private System.Windows.Forms.Timer _calendarTimer;
    private int _totalImported;
    private DateTime? _lastPoll;
    private string? _lastError;
    private string _header = "AulaSync";
    private bool _fetching;
    private System.Net.HttpListener? _httpListener;
    private const int IcsPort = 9876;

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

    private static readonly Dictionary<string, int> CalendarIntervals = new()
    {
        ["Hver 4. time"] = 4 * 3_600_000,
        ["Hver 6. time"] = 6 * 3_600_000,
        ["Hver 12. time"] = 12 * 3_600_000,
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

        LoadSubscriptions();
        _pollTimer = new System.Windows.Forms.Timer();
        _pollTimer.Interval = LoadConfig().TryGetValue("poll_interval", out var pi) ? (int)pi : 300_000;
        _pollTimer.Tick += async (_, _) => await DoFetchAsync();

        // Rullende kalender-synk (interval = total cyklus / antal subscriptions, min 30 min)
        _calendarTimer = new System.Windows.Forms.Timer();
        _calendarTimer.Interval = 1_800_000; // Beregnes ved start
        _calendarTimer.Tick += async (_, _) => await SyncNextCalendarAsync();
    }

    public async Task RunLoopAsync()
    {
        while (true)
        {
            // Silent mode: vent på bruger-klik før login
            if (Program.Silent)
            {
                _tray.Text = "AulaSync - klik for at logge ind";
                Log("Silent mode - venter på bruger-klik");
                var clicked = new TaskCompletionSource();
                void handler(object? s, MouseEventArgs e) { clicked.TrySetResult(); }
                _tray.MouseClick += handler;
                await clicked.Task;
                _tray.MouseClick -= handler;
                Program.Silent = false; // Kun silent ved første start
            }

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

            // Start ICS-webserver
            StartIcsServer();

            // Sæt menuen med det samme
            _tray.ContextMenuStrip = BuildMenu();
            ShowBalloon($"Logget ind som {_api.UserName}");
            Log($"Logget ind som {_api.UserName}");

            // Hent beskeder i baggrunden (blokerer ikke UI)
            _ = DoFetchAsync();

            // Start polling
            _pollTimer.Start();

            // Start rullende kalender-synk forskudt 5 min
            if (_subscribedEmployees.Count > 0)
            {
                // Fordel synk jævnt: total cyklus (default 6t) / antal subscriptions, min 30 min
                var totalCycleMs = LoadConfig().TryGetValue("calendar_interval", out var ci) ? (int)ci : 6 * 3_600_000;
                _calendarTimer.Interval = Math.Max(totalCycleMs / _subscribedEmployees.Count, 1_800_000);
                Log($"Kalender-sync: {_subscribedEmployees.Count} skemaer, interval {_calendarTimer.Interval / 60000} min");

                var calStartDelay = new System.Windows.Forms.Timer { Interval = 300_000 }; // 5 min
                calStartDelay.Tick += (_, _) =>
                {
                    calStartDelay.Stop();
                    calStartDelay.Dispose();
                    _calendarTimer.Start();
                    _ = SyncNextCalendarAsync();
                };
                calStartDelay.Start();
            }

            // Vent på logout/quit
            _currentWait = new TaskCompletionSource();
            await _currentWait.Task;
            _currentWait = null;

            _pollTimer.Stop();
            _calendarTimer.Stop();
            StopIcsServer();

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
            _cachedEmployees = null;
            _syncIndex = 0;
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
        menu.Items.Add("Synkronisér beskeder", null, async (_, _) => await DoFetchAsync());

        // Medarbejdere submenu
        var empMenu = new ToolStripMenuItem("Medarbejderskemaer til Outlook");
        empMenu.DropDownItems.Add("Henter...");
        empMenu.DropDownOpening += async (_, _) =>
        {
            if (empMenu.DropDownItems.Count == 1 && empMenu.DropDownItems[0].Text == "Henter...")
            {
                if (_api == null) return;
                try
                {
                    var employees = await _api.GetEmployeesAsync();
                    empMenu.DropDownItems.Clear();

                    // Lærere+pædagoger først, derefter resten
                    var teaching = employees
                        .Where(e => e.Role is "teacher" or "preschool-teacher")
                        .OrderBy(e => e.Initials).ToList();
                    var other = employees
                        .Where(e => e.Role is not "teacher" and not "preschool-teacher")
                        .OrderBy(e => e.Initials).ToList();

                    foreach (var emp in teaching)
                    {
                        var c = emp;
                        var item = new ToolStripMenuItem($"{emp.Initials} - {emp.Name}");
                        item.Click += async (_, _) => await SaveAndOpenIcs(c);
                        empMenu.DropDownItems.Add(item);
                    }

                    if (other.Count > 0)
                    {
                        empMenu.DropDownItems.Add(new ToolStripSeparator());
                        var hdr = new ToolStripMenuItem("Øvrige") { Enabled = false };
                        hdr.Font = new Font(hdr.Font, FontStyle.Bold);
                        empMenu.DropDownItems.Add(hdr);

                        foreach (var emp in other)
                        {
                            var c = emp;
                            var item = new ToolStripMenuItem($"{emp.Initials} - {emp.Name}");
                            item.Click += async (_, _) => await SaveAndOpenIcs(c);
                            empMenu.DropDownItems.Add(item);
                        }
                    }
                }
                catch (Exception ex)
                {
                    empMenu.DropDownItems.Clear();
                    empMenu.DropDownItems.Add($"Fejl: {ex.Message}");
                }
            }
        };
        menu.Items.Add(empMenu);

        // Klasseskemaer submenu
        var classMenu = new ToolStripMenuItem("Klasseskemaer til Outlook");
        classMenu.DropDownItems.Add("Henter...");
        classMenu.DropDownOpening += async (_, _) =>
        {
            if (classMenu.DropDownItems.Count == 1 && classMenu.DropDownItems[0].Text == "Henter...")
            {
                if (_api == null) return;
                try
                {
                    var groups = await _api.GetGroupsAsync();
                    classMenu.DropDownItems.Clear();
                    // Kun hovedgrupper (klasser)
                    foreach (var g in groups.Where(g => g.Type == "Hovedgruppe"))
                    {
                        var captured = g;
                        var item = new ToolStripMenuItem(g.Name);
                        item.Click += async (_, _) => await FetchAndOpenCalendar(
                            g.Name, () => FetchGroupCalendarInChunks(captured.Id));
                        classMenu.DropDownItems.Add(item);
                    }
                }
                catch (Exception ex) { classMenu.DropDownItems.Clear(); classMenu.DropDownItems.Add($"Fejl: {ex.Message}"); }
            }
        };
        menu.Items.Add(classMenu);

        // Lokaleskemaer submenu
        var roomMenu = new ToolStripMenuItem("Lokaleskemaer til Outlook");
        roomMenu.DropDownItems.Add("Henter...");
        roomMenu.DropDownOpening += async (_, _) =>
        {
            if (roomMenu.DropDownItems.Count == 1 && roomMenu.DropDownItems[0].Text == "Henter...")
            {
                if (_api == null) return;
                try
                {
                    var resources = await _api.GetResourcesAsync();
                    roomMenu.DropDownItems.Clear();
                    foreach (var r in resources)
                    {
                        var captured = r;
                        var item = new ToolStripMenuItem(r.Name);
                        item.Click += async (_, _) => await FetchAndOpenCalendar(
                            r.Name, () => FetchResourceCalendarInChunks(captured.Id));
                        roomMenu.DropDownItems.Add(item);
                    }
                }
                catch (Exception ex) { roomMenu.DropDownItems.Clear(); roomMenu.DropDownItems.Add($"Fejl: {ex.Message}"); }
            }
        };
        menu.Items.Add(roomMenu);

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

        var calInterval = new ToolStripMenuItem("Skema-synkronisering");
        foreach (var (label, ms) in CalendarIntervals)
        {
            var c2 = ms;
            var item2 = new ToolStripMenuItem(label, null, (_, _) =>
            {
                _calendarTimer.Interval = c2;
                SaveConfigValue("calendar_interval", c2);
            });
            calInterval.DropDownOpening += (_, _) => item2.Checked = _calendarTimer.Interval == c2;
            calInterval.DropDownItems.Add(item2);
        }
        settings.DropDownItems.Add(calInterval);

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
        var about = new ToolStripMenuItem("AulaSync v2.2.2") { Enabled = false };
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
        if (_api == null || _fetching) return;
        _fetching = true;
        try
        {
            await _api.TestSessionAsync();
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
        finally
        {
            _fetching = false;
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

    private const int CalendarDaysBack = 90;  // 3 måneder
    private const int CalendarDaysFwd = 90;   // 3 måneder
    private const int CalendarChunkDays = 42;  // Max per API-kald

    private async Task SaveAndOpenIcs(AulaApi.Employee emp)
    {
        await FetchAndOpenCalendar(
            $"{emp.Initials} - {emp.Name}",
            () => FetchCalendarInChunks(emp.Id),
            emp.Initials);

        // Tilføj til auto-sync subscriptions
        if (_subscribedEmployees.Add(emp.Id))
        {
            SaveSubscriptions();
            // Genberegn sync-interval
            var totalCycleMs = LoadConfig().TryGetValue("calendar_interval", out var ci) ? (int)ci : 6 * 3_600_000;
            _calendarTimer.Interval = Math.Max(totalCycleMs / _subscribedEmployees.Count, 1_800_000);
            if (!_calendarTimer.Enabled) _calendarTimer.Start();
            Log($"Subscription tilføjet: {emp.Initials} (total: {_subscribedEmployees.Count})");
        }
    }

    private async Task<List<AulaApi.CalendarEvent>> FetchCalendarInChunks(string profileId)
    {
        var all = new List<AulaApi.CalendarEvent>();
        var start = DateTime.Now.AddDays(-CalendarDaysBack);
        var end = DateTime.Now.AddDays(CalendarDaysFwd);
        for (var s = start; s < end; s = s.AddDays(CalendarChunkDays))
        {
            var e = s.AddDays(CalendarChunkDays);
            if (e > end) e = end;
            all.AddRange(await _api!.GetCalendarAsync(profileId, s, e));
        }
        return all;
    }

    private async Task<List<AulaApi.CalendarEvent>> FetchGroupCalendarInChunks(string groupId)
    {
        var all = new List<AulaApi.CalendarEvent>();
        var start = DateTime.Now.AddDays(-CalendarDaysBack);
        var end = DateTime.Now.AddDays(CalendarDaysFwd);
        for (var s = start; s < end; s = s.AddDays(CalendarChunkDays))
        {
            var e = s.AddDays(CalendarChunkDays);
            if (e > end) e = end;
            all.AddRange(await _api!.GetGroupCalendarAsync(groupId, s, e));
        }
        return all;
    }

    private async Task<List<AulaApi.CalendarEvent>> FetchResourceCalendarInChunks(string resourceId)
    {
        var all = new List<AulaApi.CalendarEvent>();
        var start = DateTime.Now.AddDays(-CalendarDaysBack);
        var end = DateTime.Now.AddDays(CalendarDaysFwd);
        for (var s = start; s < end; s = s.AddDays(CalendarChunkDays))
        {
            var e = s.AddDays(CalendarChunkDays);
            if (e > end) e = end;
            all.AddRange(await _api!.GetResourceCalendarAsync(resourceId, s, e));
        }
        return all;
    }

    private static readonly string CalendarDir = Path.Combine(ConfigDir, "kalendere");

    private Task FetchAndOpenCalendar(string calendarName, Func<Task<List<AulaApi.CalendarEvent>>> fetchFunc,
        string initials = "")
        => FetchAndSaveIcs(calendarName, fetchFunc, initials, silent: false);

    private async Task FetchAndSaveIcs(string calendarName, Func<Task<List<AulaApi.CalendarEvent>>> fetchFunc,
        string initials, bool silent)
    {
        try
        {
            // Tjek session
            try { await _api!.TestSessionAsync(); }
            catch (UnauthorizedAccessException)
            {
                ShowBalloon("Session udløbet - log ind igen");
                _quitRequested = false;
                _currentWait?.TrySetResult();
                return;
            }

            _tray.Text = $"AulaSync - henter {calendarName}...";
            var events = await fetchFunc();
            Log($"Hentet {events.Count} events for {calendarName}");

            if (events.Count == 0)
            {
                if (!silent) _tray.Text = $"AulaSync - ingen skema for {calendarName}";
                return;
            }

            // Generér ICS-fil
            var ics = AulaApi.EventsToIcs(events, calendarName, initials);
            Directory.CreateDirectory(CalendarDir);
            var safeFileName = string.Join("_", calendarName.Split(Path.GetInvalidFileNameChars()));
            var icsPath = Path.Combine(CalendarDir, $"{safeFileName}.ics");

            // Atomisk skrivning
            var tmp = icsPath + ".tmp";
            await File.WriteAllTextAsync(tmp, ics, System.Text.Encoding.UTF8);
            File.Move(tmp, icsPath, true);

            var icsUrl = $"http://localhost:{IcsPort}/{Uri.EscapeDataString(safeFileName)}.ics";
            Log($"ICS gemt: {icsPath} ({events.Count} events) → {icsUrl}");

            if (!silent)
            {
                // Abonnér via webcal:// — Outlook håndterer subscription
                var webcalUrl = $"webcal://localhost:{IcsPort}/{Uri.EscapeDataString(safeFileName)}.ics";
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                { FileName = webcalUrl, UseShellExecute = true });
                _tray.Text = $"AulaSync - {_api?.UserName}";
            }
        }
        catch (Exception ex)
        {
            if (!silent) ShowBalloon($"Fejl: {ex.Message}");
            Log($"Kalender-fejl: {ex}");
        }
    }

    private void StartIcsServer()
    {
        try
        {
            _httpListener = new System.Net.HttpListener();
            _httpListener.Prefixes.Add($"http://localhost:{IcsPort}/");
            _httpListener.Start();
            _ = Task.Run(async () =>
            {
                while (_httpListener?.IsListening == true)
                {
                    try
                    {
                        var ctx = await _httpListener.GetContextAsync();
                        var fileName = Uri.UnescapeDataString(ctx.Request.Url!.AbsolutePath.TrimStart('/'));
                        var filePath = Path.GetFullPath(Path.Combine(CalendarDir, fileName));

                        // Sikkerhed: forhindre path traversal
                        if (filePath.StartsWith(Path.GetFullPath(CalendarDir), StringComparison.OrdinalIgnoreCase)
                            && File.Exists(filePath) && filePath.EndsWith(".ics", StringComparison.OrdinalIgnoreCase))
                        {
                            var data = await File.ReadAllBytesAsync(filePath);
                            ctx.Response.ContentType = "text/calendar; charset=utf-8";
                            ctx.Response.ContentLength64 = data.Length;
                            await ctx.Response.OutputStream.WriteAsync(data);
                        }
                        else
                        {
                            ctx.Response.StatusCode = 404;
                        }
                        ctx.Response.Close();
                    }
                    catch (ObjectDisposedException) { break; }
                    catch (Exception ex) { Log($"ICS-server fejl: {ex.Message}"); }
                }
            });
            Log($"ICS-server startet på port {IcsPort}");
        }
        catch (Exception ex)
        {
            Log($"ICS-server kunne ikke starte: {ex.Message}");
        }
    }

    private void StopIcsServer()
    {
        try { _httpListener?.Stop(); _httpListener?.Close(); } catch { }
        _httpListener = null;
    }

    private void ShowBalloon(string text)
    {
        _tray.ShowBalloonTip(3000, "AulaSync", text, ToolTipIcon.Info);
    }

    private HashSet<string> _subscribedEmployees = new();

    private void LoadSubscriptions()
    {
        try
        {
            var file = Path.Combine(ConfigDir, "subscribed_calendars.json");
            if (File.Exists(file))
                _subscribedEmployees = JsonSerializer.Deserialize<HashSet<string>>(File.ReadAllText(file)) ?? new();
        }
        catch { }
    }

    private void SaveSubscriptions()
    {
        try
        {
            Directory.CreateDirectory(ConfigDir);
            File.WriteAllText(Path.Combine(ConfigDir, "subscribed_calendars.json"),
                JsonSerializer.Serialize(_subscribedEmployees, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch { }
    }

    private List<AulaApi.Employee>? _cachedEmployees;
    private int _syncIndex;

    private async Task SyncNextCalendarAsync()
    {
        if (_api == null || _subscribedEmployees.Count == 0) return;

        // Hent medarbejdere én gang, genbrug derefter
        if (_cachedEmployees == null)
        {
            try { _cachedEmployees = await _api.GetEmployeesAsync(); }
            catch { return; }
        }

        var empById = _cachedEmployees.ToDictionary(e => e.Id);
        var subList = _subscribedEmployees.ToList();
        if (subList.Count == 0) return;

        // Rullende: synk én ad gangen
        _syncIndex = _syncIndex % subList.Count;
        var empId = subList[_syncIndex];
        _syncIndex++;

        if (!empById.TryGetValue(empId, out var emp)) return;

        try
        {
            Log($"Rullende sync: {emp.Initials} ({_syncIndex}/{subList.Count})");
            await FetchAndSaveIcs(
                $"{emp.Initials} - {emp.Name}",
                () => FetchCalendarInChunks(emp.Id),
                emp.Initials, silent: true);
        }
        catch (Exception ex) { Log($"Sync fejl for {emp.Initials}: {ex.Message}"); }
    }

    private static void Log(string message)
    {
        try
        {
            Directory.CreateDirectory(ConfigDir);
            // Log-rotation: max 1 MB
            if (File.Exists(LogFile) && new FileInfo(LogFile).Length > 1_000_000)
            {
                var backup = LogFile + ".old";
                if (File.Exists(backup)) File.Delete(backup);
                File.Move(LogFile, backup);
            }
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
                key.SetValue(AutostartName, $"\"{Application.ExecutablePath}\" --silent");
            else
                key.DeleteValue(AutostartName, false);
        }
        catch { }
    }
}

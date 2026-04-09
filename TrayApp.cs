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

    private const string CalSuffix = " (Aula)";  // ASCII — til ICS X-WR-CALNAME
    private const string CalDisplay = " \u25B2"; // ▲ — Unicode display-navn via COM
    private const string AppVersion = "2.3.0";
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
            _tray.Text = $"AulaSync v{AppVersion} - {_api.UserName}";

            // Tjek for opdatering i baggrunden
            _ = Task.Run(async () =>
            {
                var newVer = await CheckForUpdateAsync();
                if (newVer != null)
                    ShowBalloon($"AulaSync v{newVer} er tilgængelig — opdatér via Indstillinger");
            });

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
            if (TotalSubscriptions > 0)
            {
                // Fordel synk jævnt: total cyklus (default 6t) / antal subscriptions, min 30 min
                RecalcCalendarInterval();
                Log($"Kalender-sync: {TotalSubscriptions} skemaer, interval {_calendarTimer.Interval / 60000} min");

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
            _cachedGroups = null;
            _cachedResources = null;
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
            if (_api == null) return;
            try
            {
                if (_cachedEmployees == null)
                    _cachedEmployees = await _api.GetEmployeesAsync();
                var employees = _cachedEmployees;

                SyncSubscriptionsWithOutlook(employees);

                empMenu.DropDownItems.Clear();

                    // Lærere+pædagoger først, derefter resten
                    var teaching = employees
                        .Where(e => e.Role is "teacher" or "preschool-teacher")
                        .OrderBy(e => e.Initials).ToList();
                    var other = employees
                        .Where(e => e.Role is not "teacher" and not "preschool-teacher")
                        .OrderBy(e => e.Initials).ToList();

                    void AddEmpItem(AulaApi.Employee emp)
                    {
                        var c = emp;
                        var isSub = _subscribedEmployees.Contains(c.Id);
                        var item = new ToolStripMenuItem($"{emp.Initials} {emp.Name}");
                        if (isSub) item.Checked = true;
                        item.Click += async (_, _) =>
                        {
                            if (_subscribedEmployees.Contains(c.Id))
                            {
                                _subscribedEmployees.Remove(c.Id);
                                CancelMoveTimer(c.Id);
                                SaveSubscriptions();
                                RecalcCalendarInterval();
                                var calName = EmpCalName(c.Initials, c.Name);
                                var safeName = string.Join("_", calName.Split(Path.GetInvalidFileNameChars()));
                                var icsFile = Path.Combine(CalendarDir, $"{safeName}.ics");
                                try { if (File.Exists(icsFile)) File.Delete(icsFile); } catch { }
                                OutlookBridge.RemoveInternetCalendar(calName);
                                OutlookBridge.RemoveInternetCalendar(calName.Replace(CalSuffix, CalDisplay));
                                ShowBalloon($"Skema fjernet: {c.Initials}");
                                Log($"Subscription fjernet: {c.Initials}");
                            }
                            else
                            {
                                await SaveAndOpenIcs(c);
                            }
                        };
                        empMenu.DropDownItems.Add(item);
                    }

                    foreach (var emp in teaching) AddEmpItem(emp);

                    if (other.Count > 0)
                    {
                        empMenu.DropDownItems.Add(new ToolStripSeparator());
                        var hdr = new ToolStripMenuItem("Øvrige") { Enabled = false };
                        hdr.Font = new Font(hdr.Font, FontStyle.Bold);
                        empMenu.DropDownItems.Add(hdr);

                        foreach (var emp in other) AddEmpItem(emp);
                    }
                }
                catch (Exception ex)
                {
                    empMenu.DropDownItems.Clear();
                    empMenu.DropDownItems.Add($"Fejl: {ex.Message}");
                }
        };
        menu.Items.Add(empMenu);

        // Klasseskemaer submenu
        var classMenu = new ToolStripMenuItem("Klasseskemaer til Outlook");
        classMenu.DropDownItems.Add("Henter...");
        classMenu.DropDownOpening += async (_, _) =>
        {
            if (_api == null) return;
            try
            {
                if (_cachedGroups == null)
                    _cachedGroups = await _api.GetGroupsAsync();
                var groups = _cachedGroups;

                classMenu.DropDownItems.Clear();
                // Kun hovedgrupper (klasser)
                foreach (var g in groups.Where(g => g.Type == "Hovedgruppe"))
                {
                    var captured = g;
                    var isSub = _subscribedGroups.ContainsKey(captured.Id);
                    var item = new ToolStripMenuItem(g.Name);
                    if (isSub) item.Checked = true;
                    item.Click += async (_, _) =>
                    {
                        if (_subscribedGroups.ContainsKey(captured.Id))
                        {
                            _subscribedGroups.Remove(captured.Id);
                            CancelMoveTimer(captured.Id);
                            SaveSubscriptions();
                            RecalcCalendarInterval();
                            var calName = CalName(captured.Name);
                            var safeName = string.Join("_", calName.Split(Path.GetInvalidFileNameChars()));
                            var icsFile = Path.Combine(CalendarDir, $"{safeName}.ics");
                            try { if (File.Exists(icsFile)) File.Delete(icsFile); } catch { }
                            OutlookBridge.RemoveInternetCalendar(calName);
                            OutlookBridge.RemoveInternetCalendar(calName.Replace(CalSuffix, CalDisplay));
                            ShowBalloon($"Skema fjernet: {captured.Name}");
                            Log($"Gruppe-subscription fjernet: {captured.Name}");
                        }
                        else
                        {
                            await FetchAndOpenCalendar(
                                CalName(captured.Name), () => FetchGroupCalendarInChunks(captured.Id),
                                empId: captured.Id);
                            if (_subscribedGroups.TryAdd(captured.Id, CalName(captured.Name)))
                            {
                                SaveSubscriptions();
                                RecalcCalendarInterval();
                                if (!_calendarTimer.Enabled) _calendarTimer.Start();
                                Log($"Gruppe-subscription tilføjet: {captured.Name} (total: {TotalSubscriptions})");
                            }
                        }
                    };
                    classMenu.DropDownItems.Add(item);
                }
            }
            catch (Exception ex) { classMenu.DropDownItems.Clear(); classMenu.DropDownItems.Add($"Fejl: {ex.Message}"); }
        };
        menu.Items.Add(classMenu);

        // Lokaleskemaer submenu
        var roomMenu = new ToolStripMenuItem("Lokaleskemaer til Outlook");
        roomMenu.DropDownItems.Add("Henter...");
        roomMenu.DropDownOpening += async (_, _) =>
        {
            if (_api == null) return;
            try
            {
                if (_cachedResources == null)
                    _cachedResources = await _api.GetResourcesAsync();
                var resources = _cachedResources;

                roomMenu.DropDownItems.Clear();
                foreach (var r in resources)
                {
                    var captured = r;
                    var isSub = _subscribedResources.ContainsKey(captured.Id);
                    var item = new ToolStripMenuItem(r.Name);
                    if (isSub) item.Checked = true;
                    item.Click += async (_, _) =>
                    {
                        if (_subscribedResources.ContainsKey(captured.Id))
                        {
                            _subscribedResources.Remove(captured.Id);
                            CancelMoveTimer(captured.Id);
                            SaveSubscriptions();
                            RecalcCalendarInterval();
                            var calName = CalName(captured.Name);
                            var safeName = string.Join("_", calName.Split(Path.GetInvalidFileNameChars()));
                            var icsFile = Path.Combine(CalendarDir, $"{safeName}.ics");
                            try { if (File.Exists(icsFile)) File.Delete(icsFile); } catch { }
                            OutlookBridge.RemoveInternetCalendar(calName);
                            OutlookBridge.RemoveInternetCalendar(calName.Replace(CalSuffix, CalDisplay));
                            ShowBalloon($"Skema fjernet: {captured.Name}");
                            Log($"Ressource-subscription fjernet: {captured.Name}");
                        }
                        else
                        {
                            await FetchAndOpenCalendar(
                                CalName(captured.Name), () => FetchResourceCalendarInChunks(captured.Id),
                                empId: captured.Id);
                            if (_subscribedResources.TryAdd(captured.Id, CalName(captured.Name)))
                            {
                                SaveSubscriptions();
                                RecalcCalendarInterval();
                                if (!_calendarTimer.Enabled) _calendarTimer.Start();
                                Log($"Ressource-subscription tilføjet: {captured.Name} (total: {TotalSubscriptions})");
                            }
                        }
                    };
                    roomMenu.DropDownItems.Add(item);
                }
            }
            catch (Exception ex) { roomMenu.DropDownItems.Clear(); roomMenu.DropDownItems.Add($"Fejl: {ex.Message}"); }
        };
        menu.Items.Add(roomMenu);

        menu.Items.Add(new ToolStripSeparator());

        // Indstillinger submenu
        var settings = new ToolStripMenuItem("Indstillinger");

        var interval = new ToolStripMenuItem("Besked-interval");
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

        var calInterval = new ToolStripMenuItem("Skema-interval");
        foreach (var (label, ms) in CalendarIntervals)
        {
            var c2 = ms;
            var item2 = new ToolStripMenuItem(label, null, (_, _) =>
            {
                SaveConfigValue("calendar_interval", c2);
                RecalcCalendarInterval();
            });
            calInterval.DropDownOpening += (_, _) =>
            {
                var saved = LoadConfig().TryGetValue("calendar_interval", out var ci) ? (int)ci : 6 * 3_600_000;
                item2.Checked = saved == c2;
            };
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
        var about = new ToolStripMenuItem($"AulaSync v{AppVersion}") { Enabled = false };
        settings.DropDownItems.Add(about);
        var updateItem = new ToolStripMenuItem("Søg efter opdatering...", null, async (_, _) =>
        {
            var newVer = await CheckForUpdateAsync();
            if (newVer != null)
            {
                var result = MessageBox.Show(
                    $"AulaSync v{newVer} er tilgængelig (du har v{AppVersion}).\n\nVil du hente den?",
                    "AulaSync - Opdatering", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (result == DialogResult.Yes)
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    { FileName = $"{ProjectUrl}/releases/latest", UseShellExecute = true });
            }
            else
            {
                ShowBalloon("Du har den nyeste version");
            }
        });
        settings.DropDownItems.Add(updateItem);

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
        var parts = new List<string>();
        parts.Add(_totalImported == 1 ? "1 ny besked" : $"{_totalImported} nye beskeder");
        if (_lastPoll.HasValue) parts.Add($"sidst tjekket {_lastPoll.Value:HH:mm}");
        var sc = TotalSubscriptions;
        if (sc > 0) parts.Add(sc == 1 ? "1 skema" : $"{sc} skemaer");
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
        catch (HttpRequestException hx) when (hx.StatusCode == System.Net.HttpStatusCode.Forbidden)
        {
            _lastError = "Session udløbet (403)";
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
        var seenIds = OutlookBridge.GetSeenThreadIds();
        var messages = new List<AulaMessage>();
        foreach (var t in threads)
        {
            var threadId = t.GetProperty("id").GetInt64().ToString();
            var isRead = t.TryGetProperty("read", out var rd) && rd.GetBoolean();

            // Skip tråde vi allerede har set og som er læste (ingen nye svar)
            if (isRead && seenIds.Contains(threadId)) continue;

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
        CancelMoveTimer(emp.Id); // Cancel evt. eksisterende move-timer
        await FetchAndOpenCalendar(
            EmpCalName(emp.Initials, emp.Name),
            () => FetchCalendarInChunks(emp.Id),
            emp.Initials,
            emp.Id);

        // Tilføj til auto-sync subscriptions
        if (_subscribedEmployees.Add(emp.Id))
        {
            SaveSubscriptions();
            RecalcCalendarInterval();
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
        string initials = "", string? empId = null)
        => FetchAndSaveIcs(calendarName, fetchFunc, initials, silent: false, empId: empId);

    private async Task FetchAndSaveIcs(string calendarName, Func<Task<List<AulaApi.CalendarEvent>>> fetchFunc,
        string initials, bool silent, string? empId = null)
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
            await File.WriteAllTextAsync(tmp, ics, new System.Text.UTF8Encoding(false));
            File.Move(tmp, icsPath, true);

            var icsUrl = $"http://localhost:{IcsPort}/{Uri.EscapeDataString(safeFileName)}.ics";
            Log($"ICS gemt: {icsPath} ({events.Count} events) → {icsUrl}");

            if (!silent)
            {
                // Tjek om kalenderen allerede findes i Outlook (undgå duplikater)
                var outlookCals = OutlookBridge.GetInternetCalendarNames();
                if (!OutlookHasCalendar(outlookCals, calendarName))
                {
                    // Abonnér via webcal:// — Outlook håndterer subscription
                    var webcalUrl = $"webcal://localhost:{IcsPort}/{Uri.EscapeDataString(safeFileName)}.ics";
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    { FileName = webcalUrl, UseShellExecute = true });

                    // Flyt til institutionsgruppe efter Outlook har tilføjet den (på UI-tråden med retry)
                    var institution = _api?.Institution ?? "AulaSync";
                    var displayName = calendarName.Replace(CalSuffix, CalDisplay);
                    var moveRetries = 0;
                    var moveTimer = new System.Windows.Forms.Timer { Interval = 10_000 };
                    moveTimer.Tick += (_, _) =>
                    {
                        moveRetries++;
                        if (OutlookBridge.MoveCalendarToGroup(calendarName, institution))
                        {
                            OutlookBridge.RenameCalendar(calendarName, displayName);
                            moveTimer.Stop();
                            moveTimer.Dispose();
                            if (empId != null) _moveTimers.Remove(empId);
                        }
                        else if (moveRetries >= 6)
                        {
                            // Prøv rename alligevel (kalenderen kan findes men ikke flyttes)
                            OutlookBridge.RenameCalendar(calendarName, displayName);
                            moveTimer.Stop();
                            moveTimer.Dispose();
                            if (empId != null) _moveTimers.Remove(empId);
                        }
                    };
                    if (empId != null) _moveTimers[empId] = moveTimer;
                    moveTimer.Start();
                }
                else
                {
                    Log($"Kalender '{calendarName}' findes allerede i Outlook — springer webcal over");
                }
                _tray.Text = $"AulaSync - {_api?.UserName}";
            }
        }
        catch (Exception ex)
        {
            ShowBalloon($"Skema-fejl: {ex.Message}");
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
                    catch (Exception) { }
                }
            });
            Log($"ICS-server startet på port {IcsPort}");
        }
        catch (Exception ex)
        {
            Log($"ICS-server kunne ikke starte: {ex.Message}");
            ShowBalloon($"Kalender-server fejl: port {IcsPort} er optaget");
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

    private static async Task<string?> CheckForUpdateAsync()
    {
        try
        {
            using var http = new HttpClient();
            http.DefaultRequestHeaders.UserAgent.ParseAdd("AulaSync");
            var json = await http.GetStringAsync($"https://api.github.com/repos/rpaasch/AulaSync/releases/latest");
            using var doc = JsonDocument.Parse(json);
            var tag = doc.RootElement.GetProperty("tag_name").GetString()?.TrimStart('v') ?? "";
            if (tag != "" && tag != AppVersion && string.Compare(tag, AppVersion, StringComparison.Ordinal) > 0)
                return tag;
        }
        catch { }
        return null;
    }

    private HashSet<string> _subscribedEmployees = new();
    private Dictionary<string, string> _subscribedGroups = new(); // id → calendarName
    private Dictionary<string, string> _subscribedResources = new(); // id → calendarName
    private readonly Dictionary<string, System.Windows.Forms.Timer> _moveTimers = new();

    private int TotalSubscriptions => _subscribedEmployees.Count + _subscribedGroups.Count + _subscribedResources.Count;

    private void LoadSubscriptions()
    {
        try
        {
            var file = Path.Combine(ConfigDir, "subscribed_calendars.json");
            if (!File.Exists(file)) return;
            var json = File.ReadAllText(file).Trim();
            if (json.StartsWith("["))
            {
                // Backward compat: old format was just an array of employee ids
                _subscribedEmployees = JsonSerializer.Deserialize<HashSet<string>>(json) ?? new();
            }
            else
            {
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;
                if (root.TryGetProperty("employees", out var emp))
                    _subscribedEmployees = JsonSerializer.Deserialize<HashSet<string>>(emp.GetRawText()) ?? new();
                if (root.TryGetProperty("groups", out var grp))
                    _subscribedGroups = JsonSerializer.Deserialize<Dictionary<string, string>>(grp.GetRawText()) ?? new();
                if (root.TryGetProperty("resources", out var res))
                    _subscribedResources = JsonSerializer.Deserialize<Dictionary<string, string>>(res.GetRawText()) ?? new();
            }
        }
        catch { }
    }

    private void SaveSubscriptions()
    {
        try
        {
            Directory.CreateDirectory(ConfigDir);
            var data = new Dictionary<string, object>
            {
                ["employees"] = _subscribedEmployees,
                ["groups"] = _subscribedGroups,
                ["resources"] = _subscribedResources
            };
            File.WriteAllText(Path.Combine(ConfigDir, "subscribed_calendars.json"),
                JsonSerializer.Serialize(data, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch { }
    }

    private List<AulaApi.Employee>? _cachedEmployees;
    private List<AulaApi.GroupInfo>? _cachedGroups;
    private List<AulaApi.ResourceInfo>? _cachedResources;
    private int _syncIndex;

    /// <summary>
    /// Fjerner subscriptions for kalendere der ikke længere findes i Outlook.
    /// Returnerer true hvis noget blev ændret.
    /// </summary>
    private bool SyncSubscriptionsWithOutlook(List<AulaApi.Employee>? employees = null)
    {
        employees ??= _cachedEmployees;

        var outlookCals = OutlookBridge.GetInternetCalendarNames();
        if (outlookCals.Count == 0)
        {
            Log("Outlook-sync sprunget over (ActiveExplorer utilgængelig)");
            return false;
        }

        var changed = false;

        // Sync employees
        if (employees != null && _subscribedEmployees.Count > 0)
        {
            var empById = employees.ToDictionary(e => e.Id);
            var removed = _subscribedEmployees
                .Where(id => !_moveTimers.ContainsKey(id)
                    && empById.TryGetValue(id, out var e)
                    && !OutlookHasCalendar(outlookCals, EmpCalName(e.Initials, e.Name)))
                .ToList();

            foreach (var id in removed)
            {
                _subscribedEmployees.Remove(id);
                CancelMoveTimer(id);
                if (empById.TryGetValue(id, out var e))
                {
                    var calName = EmpCalName(e.Initials, e.Name);
                    var safeName = string.Join("_", calName.Split(Path.GetInvalidFileNameChars()));
                    try { File.Delete(Path.Combine(CalendarDir, $"{safeName}.ics")); } catch { }
                    Log($"Subscription fjernet (ikke i Outlook): {e.Initials}");
                }
                changed = true;
            }
        }

        // Sync groups
        if (_subscribedGroups.Count > 0)
        {
            var removedGroups = _subscribedGroups
                .Where(kv => !_moveTimers.ContainsKey(kv.Key)
                    && !OutlookHasCalendar(outlookCals, kv.Value))
                .Select(kv => kv.Key).ToList();

            foreach (var id in removedGroups)
            {
                var name = _subscribedGroups[id];
                _subscribedGroups.Remove(id);
                CancelMoveTimer(id);
                var safeName = string.Join("_", name.Split(Path.GetInvalidFileNameChars()));
                try { File.Delete(Path.Combine(CalendarDir, $"{safeName}.ics")); } catch { }
                Log($"Gruppe-subscription fjernet (ikke i Outlook): {name}");
                changed = true;
            }
        }

        // Sync resources
        if (_subscribedResources.Count > 0)
        {
            var removedResources = _subscribedResources
                .Where(kv => !_moveTimers.ContainsKey(kv.Key)
                    && !OutlookHasCalendar(outlookCals, kv.Value))
                .Select(kv => kv.Key).ToList();

            foreach (var id in removedResources)
            {
                var name = _subscribedResources[id];
                _subscribedResources.Remove(id);
                CancelMoveTimer(id);
                var safeName = string.Join("_", name.Split(Path.GetInvalidFileNameChars()));
                try { File.Delete(Path.Combine(CalendarDir, $"{safeName}.ics")); } catch { }
                Log($"Ressource-subscription fjernet (ikke i Outlook): {name}");
                changed = true;
            }
        }

        if (changed)
        {
            SaveSubscriptions();
            RecalcCalendarInterval();
        }
        return changed;
    }

    private void CancelMoveTimer(string empId)
    {
        if (_moveTimers.Remove(empId, out var timer))
        {
            timer.Stop();
            timer.Dispose();
        }
    }

    /// <summary>
    /// Matcher kalendernavne — håndterer Outlook's "(N)" suffix ved duplikater.
    /// "AV - Navn" matcher "AV - Navn", "AV - Navn (1)", "AV - Navn (2)" osv.
    /// </summary>
    /// <summary>Bygger kalendernavn med 🅐 suffix for medarbejdere.</summary>
    private static string EmpCalName(string initials, string name) => $"{initials} {name}{CalSuffix}";

    /// <summary>Bygger kalendernavn med 🅐 suffix for klasser/lokaler.</summary>
    private static string CalName(string name) => $"{name}{CalSuffix}";

    private static bool OutlookHasCalendar(HashSet<string> outlookCals, string calendarName)
    {
        var displayName = calendarName.Replace(CalSuffix, CalDisplay);
        return outlookCals.Any(c => c == calendarName || c == displayName
            || MatchesWithSuffix(c, calendarName) || MatchesWithSuffix(c, displayName));
    }

    private static bool MatchesWithSuffix(string outlookName, string expected)
        => outlookName.StartsWith(expected) && outlookName.Length > expected.Length
            && outlookName[expected.Length] == ' ' && outlookName[expected.Length + 1] == '(';

    private async Task SyncNextCalendarAsync()
    {
        if (_api == null || TotalSubscriptions == 0) return;

        // Hent medarbejdere én gang, genbrug derefter
        if (_cachedEmployees == null && _subscribedEmployees.Count > 0)
        {
            try { _cachedEmployees = await _api.GetEmployeesAsync(); }
            catch { return; }
        }

        SyncSubscriptionsWithOutlook();

        // Byg unified liste: (calendarName, initials, fetchFunc)
        var syncList = new List<(string calName, string initials, Func<Task<List<AulaApi.CalendarEvent>>> fetch)>();

        if (_cachedEmployees != null)
        {
            var empById = _cachedEmployees.ToDictionary(e => e.Id);
            foreach (var id in _subscribedEmployees)
            {
                if (empById.TryGetValue(id, out var emp))
                    syncList.Add((EmpCalName(emp.Initials, emp.Name), emp.Initials, () => FetchCalendarInChunks(id)));
            }
        }

        foreach (var (id, name) in _subscribedGroups)
        {
            var capturedId = id;
            syncList.Add((name, "", () => FetchGroupCalendarInChunks(capturedId)));
        }

        foreach (var (id, name) in _subscribedResources)
        {
            var capturedId = id;
            syncList.Add((name, "", () => FetchResourceCalendarInChunks(capturedId)));
        }

        if (syncList.Count == 0) return;

        _syncIndex = _syncIndex % syncList.Count;
        var (calName, initials, fetchFunc) = syncList[_syncIndex];
        _syncIndex++;

        try
        {
            Log($"Rullende sync: {calName} ({_syncIndex}/{syncList.Count})");
            await FetchAndSaveIcs(calName, fetchFunc, initials, silent: true);

            // Sørg for kalenderen er i institutionsgruppen + har ▲ display-navn
            var institution = _api?.Institution ?? "AulaSync";
            OutlookBridge.MoveCalendarToGroup(calName, institution);
            var displayName = calName.Replace(CalSuffix, CalDisplay);
            OutlookBridge.RenameCalendar(calName, displayName);
        }
        catch (Exception ex)
        {
            ShowBalloon($"Sync-fejl: {calName} - {ex.Message}");
            Log($"Sync fejl for {calName}: {ex.Message}");
        }
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

    private void RecalcCalendarInterval()
    {
        var totalCycleMs = LoadConfig().TryGetValue("calendar_interval", out var ci) ? (int)ci : 6 * 3_600_000;
        var count = Math.Max(TotalSubscriptions, 1);
        _calendarTimer.Interval = Math.Max(totalCycleMs / count, 1_800_000);
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

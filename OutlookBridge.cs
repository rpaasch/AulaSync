using System.Runtime.InteropServices;
using System.Text.Json;
using System.Globalization;

namespace AulaSync;

/// <summary>
/// Opretter Aula-beskeder i Outlook og sætter korrekte tidsstempler via direkte MAPI.
/// </summary>
class OutlookBridge
{
    [DllImport("mapi32.dll")] static extern int MAPIInitialize(IntPtr lpMapiInit);
    [DllImport("mapi32.dll")] static extern int MAPILogonEx(IntPtr ulUIParam,
        [MarshalAs(UnmanagedType.LPWStr)] string? profile,
        [MarshalAs(UnmanagedType.LPWStr)] string? password,
        uint flags, out IntPtr session);
    [DllImport("mapi32.dll")] static extern void MAPIUninitialize();

    const uint MAPI_EXTENDED = 0x20, MAPI_USE_DEFAULT = 0x40, MAPI_NO_MAIL = 0x8000;
    const uint MAPI_UNICODE = 0x80000000, MDB_WRITE = 0x04, MAPI_DEFERRED_ERRORS = 0x08;
    const uint MAPI_BEST_ACCESS = 0x10;
    const uint PR_MESSAGE_DELIVERY_TIME = 0x0E060040;
    const uint PR_CLIENT_SUBMIT_TIME = 0x00390040;
    const uint PR_MESSAGE_FLAGS = 0x0E070003;
    const uint PR_SENDER_NAME_W = 0x0C1A001F;
    const uint PR_SENT_REPRESENTING_NAME_W = 0x0042001F;
    const int MSGFLAG_READ = 0x01;

    delegate int OpenMsgStoreDel(IntPtr self, IntPtr ui, uint cb, IntPtr eid, IntPtr iface, uint flags, out IntPtr store);
    delegate int OpenEntryDel(IntPtr self, uint cb, IntPtr eid, IntPtr iface, uint flags, out uint type, out IntPtr obj);
    delegate int SetPropsDel(IntPtr self, uint count, IntPtr props, out IntPtr problems);
    delegate int SaveChangesDel(IntPtr self, uint flags);

    static T V<T>(IntPtr obj, int slot) where T : Delegate =>
        Marshal.GetDelegateForFunctionPointer<T>(Marshal.ReadIntPtr(Marshal.ReadIntPtr(obj), slot * IntPtr.Size));

    static int SetOneProp(IntPtr pMsg, uint tag, long value)
    {
        IntPtr p = Marshal.AllocHGlobal(16);
        unsafe { byte* b = (byte*)p; *(uint*)b = tag; *(uint*)(b + 4) = 0; *(long*)(b + 8) = value; }
        int hr = V<SetPropsDel>(pMsg, 8)(pMsg, 1, p, out _);
        Marshal.FreeHGlobal(p);
        return hr;
    }

    private const string AulaFolderName = "Aula";
    private const string SeenFileName = "seen_threads.json";

    private static string ConfigDir => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".aulasync");
    private static string SeenFilePath => Path.Combine(ConfigDir, SeenFileName);

    private static void Log(string msg)
    {
        try
        {
            var logFile = Path.Combine(ConfigDir, "aulasync.log");
            File.AppendAllText(logFile, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [Outlook] {msg}\n");
        }
        catch { }
    }

    public static HashSet<string> GetSeenThreadIds()
    {
        var seen = LoadSeen();
        var ids = new HashSet<string>();
        foreach (var key in seen.Keys)
        {
            var sep = key.IndexOf('_');
            if (sep > 0) ids.Add(key[..sep]);
        }
        return ids;
    }

    private static Dictionary<string, JsonElement> LoadSeen()
    {
        Directory.CreateDirectory(ConfigDir);
        if (!File.Exists(SeenFilePath)) return new();
        try
        {
            var json = File.ReadAllText(SeenFilePath);
            return JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(json) ?? new();
        }
        catch { return new(); }
    }

    private static readonly object _seenLock = new();

    private static void SaveSeen(Dictionary<string, JsonElement> seen)
    {
        lock (_seenLock)
        {
            Directory.CreateDirectory(ConfigDir);
            // Atomisk skrivning via temp-fil
            var tmp = SeenFilePath + ".tmp";
            File.WriteAllText(tmp, JsonSerializer.Serialize(seen, new JsonSerializerOptions { WriteIndented = true }));
            File.Move(tmp, SeenFilePath, true);
        }
    }

    public static int CreateOutlookItems(List<AulaMessage> messages)
    {
        if (messages.Count == 0) return 0;

        var seen = LoadSeen();
        int newCount = 0;

        // Outlook OOM
        dynamic outlook = Activator.CreateInstance(Type.GetTypeFromProgID("Outlook.Application")!)!;
        dynamic ns = outlook.GetNamespace("MAPI");

        // Altid "Aula"-undermappe i Indbakken
        dynamic aulaFolder = GetOrCreateAulaFolder(ns);

        string storeId = aulaFolder.StoreID;

        // MAPI session til timestamps
        MAPIInitialize(IntPtr.Zero);
        MAPILogonEx(IntPtr.Zero, null, null,
            MAPI_EXTENDED | MAPI_USE_DEFAULT | MAPI_NO_MAIL | MAPI_UNICODE,
            out IntPtr pSession);

        // Åbn store via MAPI
        byte[] storeEid = Convert.FromHexString(storeId);
        IntPtr pStoreEid = Marshal.AllocHGlobal(storeEid.Length);
        Marshal.Copy(storeEid, 0, pStoreEid, storeEid.Length);
        V<OpenMsgStoreDel>(pSession, 5)(pSession, IntPtr.Zero,
            (uint)storeEid.Length, pStoreEid, IntPtr.Zero,
            MDB_WRITE | MAPI_DEFERRED_ERRORS, out IntPtr pStore);
        Marshal.FreeHGlobal(pStoreEid);

        foreach (var msg in messages)
        {
            var seenKey = $"{msg.ThreadId}_{msg.Timestamp}";
            if (seen.ContainsKey(seenKey))
            {
                Log($"Skip (seen): {seenKey}");
                continue;
            }
            Log($"NY besked: tråd {msg.ThreadId}");

            try
            {
                // Opret som PostItem (read-only i undermappe)
                dynamic mail = outlook.CreateItem(6); // PostItem
                mail.Subject = $"[Aula] [{msg.Sender}] {msg.Subject}";
                mail.HTMLBody = FormatHtml(msg);
                mail.UnRead = !msg.IsRead;
                mail.Importance = msg.IsRead ? 1 : 2;

                mail.Save();
                mail.Move(aulaFolder);

                // Find beskeden i mappen (senest oprettet)
                dynamic items = aulaFolder.Items;
                items.Sort("[ReceivedTime]", true);
                dynamic latest = items.GetFirst();
                string entryId = latest.EntryID;

                // Patch timestamps via MAPI
                if (DateTime.TryParse(msg.Timestamp, CultureInfo.InvariantCulture,
                    DateTimeStyles.RoundtripKind, out var dt))
                {
                    PatchTimestamp(pStore, entryId, dt, msg.Sender);
                }

                seen[seenKey] = JsonSerializer.SerializeToElement(new
                {
                    thread_id = msg.ThreadId,
                    imported_at = DateTime.Now.ToString("o"),
                });
                newCount++;
                SaveSeen(seen); // Gem efter hver besked så vi ikke mister progress
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Fejl ved import af tråd {msg.ThreadId}: {ex.Message}");
            }
        }

        SaveSeen(seen);
        try { if (pStore != IntPtr.Zero) Marshal.Release(pStore); } catch { }
        try { if (pSession != IntPtr.Zero) Marshal.Release(pSession); } catch { }
        try { MAPIUninitialize(); } catch { }

        return newCount;
    }

    private static void PatchTimestamp(IntPtr pStore, string entryIdHex, DateTime timestamp, string sender)
    {
        if (string.IsNullOrEmpty(entryIdHex)) return;

        byte[] msgEid;
        try { msgEid = Convert.FromHexString(entryIdHex); }
        catch { return; }

        IntPtr pMsgEid = Marshal.AllocHGlobal(msgEid.Length);
        Marshal.Copy(msgEid, 0, pMsgEid, msgEid.Length);

        int hr = V<OpenEntryDel>(pStore, 17)(pStore,
            (uint)msgEid.Length, pMsgEid, IntPtr.Zero,
            MAPI_BEST_ACCESS, out _, out IntPtr pMsg);
        Marshal.FreeHGlobal(pMsgEid);
        if (hr != 0) return;

        try
        {
            long ft = timestamp.ToUniversalTime().ToFileTimeUtc();

            SetOneProp(pMsg, PR_MESSAGE_FLAGS, MSGFLAG_READ);

            IntPtr pSender = Marshal.StringToHGlobalUni(sender ?? "");
            SetOneProp(pMsg, PR_SENDER_NAME_W, pSender.ToInt64());
            SetOneProp(pMsg, PR_SENT_REPRESENTING_NAME_W, pSender.ToInt64());
            Marshal.FreeHGlobal(pSender);

            SetOneProp(pMsg, PR_MESSAGE_DELIVERY_TIME, ft);
            SetOneProp(pMsg, PR_CLIENT_SUBMIT_TIME, ft);

            V<SaveChangesDel>(pMsg, 4)(pMsg, 0);
        }
        finally
        {
            Marshal.Release(pMsg);
        }
    }

    private static dynamic GetOrCreateAulaFolder(dynamic ns)
    {
        dynamic inbox = ns.GetDefaultFolder(6);
        foreach (dynamic f in inbox.Folders)
            if (f.Name == AulaFolderName) return f;
        Log($"Opretter mappe '{AulaFolderName}' i Indbakken");
        return inbox.Folders.Add(AulaFolderName);
    }

    public static void ResetSeen()
    {
        if (File.Exists(SeenFilePath)) File.Delete(SeenFilePath);
    }

    public static HashSet<string> GetInternetCalendarNames()
    {
        var names = new HashSet<string>();
        try
        {
            dynamic outlook = Activator.CreateInstance(Type.GetTypeFromProgID("Outlook.Application")!)!;
            dynamic explorer = outlook.ActiveExplorer();
            if (explorer == null)
            {
                Log("GetInternetCalendarNames: ActiveExplorer er null");
                return names;
            }
            dynamic calModule = explorer.NavigationPane.Modules.GetNavigationModule(1);
            dynamic navGroups = calModule.NavigationGroups;
            for (int g = 1; g <= navGroups.Count; g++)
            {
                var grp = navGroups[g];
                for (int f = 1; f <= grp.NavigationFolders.Count; f++)
                {
                    try { names.Add((string)grp.NavigationFolders[f].DisplayName); } catch { }
                }
            }
            Log($"Outlook kalendere ({names.Count}): {string.Join(", ", names)}");
        }
        catch (Exception ex) { Log($"GetInternetCalendarNames fejl: {ex.Message}"); }
        return names;
    }

    public static bool MoveCalendarToGroup(string calendarName, string groupName)
    {
        try
        {
            dynamic outlook = Activator.CreateInstance(Type.GetTypeFromProgID("Outlook.Application")!)!;
            dynamic explorer = outlook.ActiveExplorer();
            if (explorer == null)
            {
                Log($"MoveCalendarToGroup: ActiveExplorer er null - kan ikke flytte '{calendarName}'");
                return false;
            }
            dynamic calModule = explorer.NavigationPane.Modules.GetNavigationModule(1); // olModuleCalendar
            dynamic navGroups = calModule.NavigationGroups;

            // Find eller opret målgruppen
            dynamic targetGroup = null;
            for (int i = 1; i <= navGroups.Count; i++)
            {
                if (navGroups[i].Name == groupName)
                { targetGroup = navGroups[i]; break; }
            }
            targetGroup ??= navGroups.Create(groupName);

            // Søg i alle eksisterende grupper efter kalenderen
            for (int g = 1; g <= navGroups.Count; g++)
            {
                var grp = navGroups[g];
                if (grp.Name == groupName) continue; // spring målgruppen over

                for (int f = grp.NavigationFolders.Count; f >= 1; f--)
                {
                    try
                    {
                        var navFolder = grp.NavigationFolders[f];
                        string displayName = navFolder.DisplayName;
                        if (MatchesCalendarName(displayName, calendarName))
                        {
                            // Hent folder-objektet og tilføj til målgruppe
                            dynamic folder = navFolder.Folder;
                            targetGroup.NavigationFolders.Add(folder);
                            // Fjern fra den gamle gruppe
                            grp.NavigationFolders.Remove(navFolder);
                            Log($"Kalender '{displayName}' flyttet til gruppe '{groupName}'");
                            return true;
                        }
                    }
                    catch { }
                }
            }
            Log($"Kalender '{calendarName}' ikke fundet i Outlook (endnu)");
            return false;
        }
        catch (Exception ex)
        {
            Log($"Kunne ikke flytte kalender til gruppe: {ex.Message}");
            return false;
        }
    }

    public static bool RemoveInternetCalendar(string name)
    {
        try
        {
            dynamic outlook = Activator.CreateInstance(Type.GetTypeFromProgID("Outlook.Application")!)!;
            dynamic explorer = outlook.ActiveExplorer();
            if (explorer == null)
            {
                Log($"RemoveInternetCalendar: ActiveExplorer er null - kan ikke fjerne '{name}'");
                return false;
            }
            dynamic calModule = explorer.NavigationPane.Modules.GetNavigationModule(1);
            dynamic navGroups = calModule.NavigationGroups;
            for (int g = 1; g <= navGroups.Count; g++)
            {
                var grp = navGroups[g];
                for (int f = grp.NavigationFolders.Count; f >= 1; f--)
                {
                    try
                    {
                        var navFolder = grp.NavigationFolders[f];
                        if (MatchesCalendarName(navFolder.DisplayName, name))
                        {
                            dynamic folder = navFolder.Folder;
                            folder.Delete();
                            Log($"Outlook-kalender fjernet: {navFolder.DisplayName}");
                            // Fortsæt for at fjerne evt. duplikater
                        }
                    }
                    catch { }
                }
            }
            return false;
        }
        catch (Exception ex)
        {
            Log($"Kunne ikke fjerne kalender: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Matcher kalendernavne — håndterer Outlook's "(N)" suffix ved duplikater.
    /// </summary>
    private static bool MatchesCalendarName(string outlookName, string expectedName)
        => outlookName == expectedName
            || (outlookName.StartsWith(expectedName) && outlookName.Length > expectedName.Length
                && outlookName[expectedName.Length] == ' ' && outlookName[expectedName.Length + 1] == '(');

    private static string FormatHtml(AulaMessage msg)
    {
        var senderRole = msg.SenderMeta.Contains(" - ") ? msg.SenderMeta.Split(" - ")[0] : msg.SenderMeta;
        var formattedTime = msg.Timestamp;
        if (DateTime.TryParse(msg.Timestamp, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var dt))
            formattedTime = dt.ToString("dd. MMM yyyy 'kl.' HH:mm", new CultureInfo("da-DK"));

        var html = $@"<html><head><meta charset='utf-8'>
<style>
body {{ font-family: Segoe UI, sans-serif; font-size: 10pt; color: #333; margin: 16px; }}
.hdr {{ background: #2c3e87; color: white; padding: 12px 16px; border-radius: 6px 6px 0 0; overflow: hidden; }}
.hdr h2 {{ margin: 0; font-size: 13pt; display: inline; }}
.hdr a {{ color: white; background: rgba(255,255,255,0.2); padding: 4px 12px; border-radius: 4px; font-size: 9pt; text-decoration: none; float: right; margin-top: 2px; }}
.meta {{ background: #f0f2f8; padding: 10px 16px; font-size: 9pt; color: #555; border-bottom: 1px solid #d0d5e5; }}
.meta b {{ color: #333; }}
.role {{ color: #888; font-style: italic; }}
.body {{ padding: 16px; background: white; border: 1px solid #e0e3ed; border-top: none; }}
.footer {{ font-size: 8pt; color: #999; margin-top: 20px; padding-top: 10px; border-top: 1px solid #eee; }}
.footer a {{ color: #2c3e87; text-decoration: none; }}
.prev {{ margin: 8px 0; padding: 10px; background: #f8f9fc; border-left: 3px solid #2c3e87; border-radius: 0 4px 4px 0; }}
.prev .sender {{ font-weight: 600; }}
.prev .time {{ color: #999; font-size: 8pt; margin-left: 8px; }}
</style></head><body>
<div class='hdr'><h2>{System.Net.WebUtility.HtmlEncode(msg.Subject)}</h2>
<a href='https://www.aula.dk/portal/#/beskeder/{msg.ThreadId}'>Åbn i Aula ↗</a></div>
<div class='meta'><b>Fra:</b> {System.Net.WebUtility.HtmlEncode(msg.Sender)}";

        if (!string.IsNullOrEmpty(senderRole))
            html += $" <span class='role'>({senderRole})</span>";
        html += $"<br><b>Dato:</b> {formattedTime}";
        if (msg.TotalRecipients > 0)
        {
            html += $"<br><b>Til:</b> {msg.TotalRecipients} modtagere";
            if (msg.TotalMessages > 1) html += $" · {msg.TotalMessages} beskeder i tråden";
        }
        html += $"</div><div class='body'>{msg.Text}</div>";

        // Tidligere beskeder
        if (msg.AllMessages.Count > 1)
        {
            html += "<div style='padding: 0 16px 16px; background: white; border: 1px solid #e0e3ed; border-top: none;'>";
            html += "<p style='color: #666; font-size: 9pt; margin-top: 16px;'><b>Tidligere i tråden:</b></p>";
            foreach (var m in msg.AllMessages.Skip(1))
            {
                var mTime = "";
                if (DateTime.TryParse(m.Timestamp, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var mdt))
                    mTime = mdt.ToString("dd. MMM 'kl.' HH:mm", new CultureInfo("da-DK"));
                var mRole = m.SenderMeta.Contains(" - ") ? m.SenderMeta.Split(" - ")[0] : "";
                html += $"<div class='prev'><span class='sender'>{System.Net.WebUtility.HtmlEncode(m.Sender)}</span>";
                if (!string.IsNullOrEmpty(mRole)) html += $" <span class='role'>({System.Net.WebUtility.HtmlEncode(mRole)})</span>";
                html += $"<span class='time'>{mTime}</span><br>{m.Text}</div>";
            }
            html += "</div>";
        }

        html += $@"<div class='footer'>
<a href='https://www.aula.dk/portal/#/beskeder/{msg.ThreadId}'>Åbn i Aula</a> · Importeret af AulaSync
</div></body></html>";

        return html;
    }
}

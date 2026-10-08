using AulaSync.Core;

namespace AulaSync.App;

public enum TrayCommand { LogIn, RetryServer, Open, Refresh, Settings, Quit }

// Command = null: deaktiveret statuslinje. Shortcut: tast med ⌘ på Mac (vises ikke på Windows).
public sealed record TrayItem(string Label, TrayCommand? Command, char? Shortcut = null)
{
    public static readonly TrayItem Separator = new("-", null);
}

// Ikonets menu (spec §3.2): statuslinje · Åbn AulaSync… · Opdatér nu · Indstillinger… · — · Afslut AulaSync.
// Ved problemer står en aktiv handling øverst (Log ind igen…); er et importeret skema ændret, åbner den hovedvinduet.
public static class TrayMenu
{
    public static IReadOnlyList<TrayItem> Build(SyncStatus status, bool serverRunning, TimeZoneInfo tz,
        IReadOnlyList<ScheduleRef>? changedImports = null)
    {
        var items = new List<TrayItem>();
        if (!status.LoggedIn) items.Add(new("Log ind igen…", TrayCommand.LogIn));
        else if (!serverRunning) items.Add(new("Start kalender-server igen", TrayCommand.RetryServer));
        else if (changedImports is { Count: > 0 }) items.Add(new($"{StatusText.ImportChanged(changedImports)}…", TrayCommand.Open));
        items.Add(new(StatusText.Menu(status, serverRunning, tz), null));
        items.Add(TrayItem.Separator);
        items.Add(new("Åbn AulaSync…", TrayCommand.Open));
        items.Add(new("Opdatér nu", TrayCommand.Refresh, 'r'));
        items.Add(new("Indstillinger…", TrayCommand.Settings, ','));
        items.Add(TrayItem.Separator);
        items.Add(new("Afslut AulaSync", TrayCommand.Quit, 'q'));
        return items;
    }

    // Tre tilstande: normal, opdaterer og kræver handling (lille prik), også når et importeret skema er ændret.
    public static TrayState IconState(SyncStatus status, bool serverRunning, bool importChanged = false) =>
        StatusText.Level(status, serverRunning) == StatusLevel.Error || importChanged ? TrayState.Attention
        : status.Progress is not null ? TrayState.Updating
        : TrayState.Normal;
}

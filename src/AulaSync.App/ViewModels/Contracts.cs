using AulaSync.Core;

namespace AulaSync.App;

// Vinduer og dialoger, som visningsmodellerne kan bede om. Implementeres af AppWindows (Task 18).
public interface IDialogs
{
    void ShowLogin();
    void ShowSettings();
    void ShowOnboarding();  // første start, også igen fra Indstillinger › Hjælp
    Task ShowAddScheduleAsync();
    Task<bool> ShowImportGuideAsync(ScheduleRef schedule);   // true = brugeren klikkede Færdig
    Task ShowOutlookFallbackAsync(ScheduleRef schedule);
    Task<bool> ConfirmLogoutAsync();
}

// Hovedknappen og "…"-menuen pr. skema. Implementeres af MainActions (Task 16).
public interface IMainActions
{
    Task RunAsync(ScheduleRef schedule);
    Task CopyAddressAsync(ScheduleRef schedule);
    void RevealFile(ScheduleRef schedule);
}

public static class CalendarChoice
{
    // Det valgte kalenderprogram. Mac falder tilbage til Apple Kalender; Windows har intet standardvalg.
    public static CalendarApp Current(ConfigStore config, bool isMac) =>
        config.Load().CalendarApp ?? CalendarApps.DefaultFor(isMac) ?? CalendarApp.Other;
}

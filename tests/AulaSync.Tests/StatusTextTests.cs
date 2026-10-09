using AulaSync.Core;
using static AulaSync.Tests.TestData;

namespace AulaSync.Tests;

public class StatusTextTests
{
    // 12:02 UTC = 14:02 i København (sommertid)
    static readonly DateTimeOffset At = new(2026, 10, 3, 12, 2, 0, TimeSpan.Zero);
    static readonly DateTimeOffset Next = At.AddHours(6);
    static SyncStatus Ok(int n = 5) => new(ConnectionState.Online, n, At, Next, null, null);

    [Fact] public void Menu_ok() => Assert.Equal("5 skemaer · opdateret 14:02", StatusText.Menu(Ok(), true, Copenhagen));
    [Fact] public void Menu_singular() => Assert.Equal("1 skema · opdateret 14:02", StatusText.Menu(Ok(1), true, Copenhagen));
    [Fact] public void Menu_logged_out() => Assert.Equal("Logget ud af Aula", StatusText.Menu(Ok() with { Connection = ConnectionState.LoggedOut }, true, Copenhagen));
    [Fact] public void Menu_offline() => Assert.Equal("Ingen forbindelse til Aula", StatusText.Menu(Ok() with { Connection = ConnectionState.Offline }, true, Copenhagen));
    [Fact] public void Menu_port_busy() => Assert.Equal("Kalender-server kunne ikke starte", StatusText.Menu(Ok(), false, Copenhagen));
    [Fact] public void Menu_progress() => Assert.Equal("Henter 2 af 5…", StatusText.Menu(Ok() with { Progress = new(2, 5) }, true, Copenhagen));

    [Fact] public void Bottom_today() => Assert.Equal("Opdateret i dag 14:02 · næste 20:02", StatusText.BottomBar(Ok(), At.AddHours(1), Copenhagen));
    [Fact] public void Bottom_yesterday() => Assert.Equal("Opdateret i går 14:02 · næste 20:02", StatusText.BottomBar(Ok(), At.AddDays(1), Copenhagen));
    [Fact] public void Bottom_older() => Assert.Equal("Opdateret 3. okt. 14:02 · næste 20:02", StatusText.BottomBar(Ok(), At.AddDays(3), Copenhagen));
    [Fact] public void Bottom_never() => Assert.Equal("Ikke opdateret endnu", StatusText.BottomBar(Ok() with { LastSuccess = null, NextSync = null }, At, Copenhagen));
    [Fact] public void Bottom_progress() => Assert.Equal("Henter 2 af 5…", StatusText.BottomBar(Ok() with { Progress = new(2, 5) }, At, Copenhagen));

    [Fact]
    public void Banners_most_important_first()
    {
        var all = Ok() with { Connection = ConnectionState.LoggedOut };
        Assert.Equal(new Banner(BannerKind.LoggedOut, "Du er logget ud af Aula. Kalenderne viser stadig de seneste skemaer.", "Log ind igen"), StatusText.Banner(all, false, 9876));
        Assert.Equal(new Banner(BannerKind.PortBusy, "Kalender-server kunne ikke starte: port 9877 er optaget af et andet program.", "Prøv igen"), StatusText.Banner(Ok(), false, 9877));
        Assert.Equal(new Banner(BannerKind.Offline, "Ingen forbindelse til Aula.", null), StatusText.Banner(Ok() with { Connection = ConnectionState.Offline }, true, 9876));
        Assert.Null(StatusText.Banner(Ok(), true, 9876));
    }

    [Fact]
    public void Levels()
    {
        Assert.Equal(StatusLevel.Ok, StatusText.Level(Ok(), true));
        Assert.Equal(StatusLevel.Warning, StatusText.Level(Ok() with { LastError = "x" }, true));
        Assert.Equal(StatusLevel.Warning, StatusText.Level(Ok() with { Connection = ConnectionState.Offline }, true));
        Assert.Equal(StatusLevel.Error, StatusText.Level(Ok() with { Connection = ConnectionState.LoggedOut }, true));
        Assert.Equal(StatusLevel.Error, StatusText.Level(Ok(), false));
    }

    static readonly ScheduleRef Anna = new(ScheduleKind.Employee, "1", "Anna Eksempel", "AE");

    [Fact] public void Row_lessons() => Assert.Equal(new RowText("Medarbejder · 214 lektioner", false), StatusText.Row(Anna, new(214, At, null, null), Next, Copenhagen));
    // Medarbejdere vises med rollen fra Aula (institutionRole); "other" eller ukendt rolle giver "Medarbejder".
    [Theory]
    [InlineData("teacher", "Lærer")]
    [InlineData("preschool-teacher", "Pædagog")]
    [InlineData("leader", "Leder")]
    [InlineData("other", "Medarbejder")]
    [InlineData("", "Medarbejder")]
    public void Row_shows_employee_role(string role, string label) =>
        Assert.Equal(new RowText($"{label} · 3 lektioner", false), StatusText.Row(Anna with { Role = role }, new(3, At, null, null), Next, Copenhagen));

    [Fact] public void Row_one_lesson() => Assert.Equal(new RowText("Lokale · 1 lektion", false), StatusText.Row(new(ScheduleKind.Resource, "4", "53"), new(1, At, null, null), Next, Copenhagen));
    [Fact] public void Row_pending() => Assert.Equal(new RowText("Klasse · henter…", false), StatusText.Row(new(ScheduleKind.Group, "3", "7A"), new(null, null, null, null), Next, Copenhagen));
    [Fact] public void Row_error() => Assert.Equal(new RowText("Kunne ikke hentes kl. 14:02 · prøver igen 20:02", true), StatusText.Row(Anna, new(214, null, "x", At), Next, Copenhagen));

    [Fact]
    public void Import_changed_names_one_schedule_or_counts_several()
    {
        Assert.Equal("7A er ændret siden import", StatusText.ImportChanged([new(ScheduleKind.Group, "3", "7A")]));
        Assert.Equal("2 skemaer er ændret siden import", StatusText.ImportChanged([new(ScheduleKind.Group, "3", "7A"), Anna]));
    }

    [Fact] public void Short_date() => Assert.Equal("3. okt.", StatusText.ShortDate(At, Copenhagen));
}

public class RowPresenterTests
{
    static readonly DateTimeOffset At = new(2026, 10, 3, 12, 0, 0, TimeSpan.Zero);
    static readonly DateTimeOffset Later = At.AddDays(3);
    static readonly ScheduleRef SevenA = new(ScheduleKind.Group, "3", "7A");

    static RowButtons Buttons(CalendarApp app, Subscription s, DateTimeOffset now, bool changed = false) =>
        RowPresenter.Buttons(app, s, changed, now, Copenhagen);

    // "✓ Tilføjet" først, når kalenderprogrammet har hentet filen; ellers venter rækken og tilbyder så knappen igen.
    [Fact]
    public void Subscribe_waits_for_the_calendar_to_fetch()
    {
        var added = new Subscription(SevenA, AddedAt: At);
        Assert.Equal(new RowButtons("Tilføj til Kalender", null, null, "Tilføj igen"), Buttons(CalendarApp.AppleCalendar, new(SevenA), Later));
        Assert.Equal(new RowButtons(null, "Venter på Kalender…", null, "Tilføj igen", Waiting: true), Buttons(CalendarApp.AppleCalendar, added, At.AddSeconds(30)));
        Assert.Equal(new RowButtons(null, "✓ Tilføjet", null, "Tilføj igen"), Buttons(CalendarApp.AppleCalendar, added with { FetchedAt = At.AddSeconds(10) }, Later));
        Assert.Equal(new RowButtons("Tilføj til Kalender", null, "Kalender hentede ikke skemaet", "Tilføj igen"),
            Buttons(CalendarApp.AppleCalendar, added, At + RowPresenter.FetchWait));
        Assert.Equal(new RowButtons(null, "Venter på Outlook…", null, "Tilføj igen", Waiting: true), Buttons(CalendarApp.OutlookClassic, added, At.AddSeconds(30)));
        Assert.Equal(new RowButtons("Tilføj til Outlook", null, "Outlook hentede ikke skemaet", "Tilføj igen"), Buttons(CalendarApp.OutlookClassic, added, Later));
    }

    [Fact]
    public void Fetched_without_click_counts_as_added() =>
        Assert.Equal(new RowButtons(null, "✓ Tilføjet", null, "Tilføj igen"), Buttons(CalendarApp.AppleCalendar, new(SevenA, FetchedAt: At), Later));

    // Rækken skifter af sig selv (fra "Venter på …" til "… hentede ikke skemaet") kun, mens den venter.
    [Fact]
    public void Changes_by_itself_only_while_waiting()
    {
        var added = new Subscription(SevenA, AddedAt: At);
        Assert.Equal(At + RowPresenter.FetchWait, RowPresenter.ChangesAt(CalendarApp.AppleCalendar, added, At.AddSeconds(30)));
        Assert.Null(RowPresenter.ChangesAt(CalendarApp.AppleCalendar, added, At + RowPresenter.FetchWait));
        Assert.Null(RowPresenter.ChangesAt(CalendarApp.AppleCalendar, added with { FetchedAt = At }, At.AddSeconds(30)));
        Assert.Null(RowPresenter.ChangesAt(CalendarApp.Other, added, At.AddSeconds(30)));
        Assert.Null(RowPresenter.ChangesAt(CalendarApp.OutlookImport, added, At.AddSeconds(30)));
        Assert.Null(RowPresenter.ChangesAt(CalendarApp.AppleCalendar, new(SevenA), At));
    }

    [Fact]
    public void Import_states()
    {
        Assert.Equal(new RowButtons("Importér…", null, null, "Importér igen…"), Buttons(CalendarApp.OutlookImport, new(SevenA), Later));
        var imported = new Subscription(SevenA, ImportedAt: At, ImportHash: "h");
        Assert.Equal(new RowButtons(null, "Importeret 3. okt.", null, "Importér igen…"), Buttons(CalendarApp.OutlookImport, imported, Later));
        Assert.Equal(new RowButtons("Importér igen…", "Importeret 3. okt.", "Ændret siden import", "Importér igen…"),
            Buttons(CalendarApp.OutlookImport, imported, Later, changed: true));
    }

    [Fact]
    public void Import_today_says_i_dag() =>
        Assert.Equal(new RowButtons(null, "Importeret i dag", null, "Importér igen…"),
            Buttons(CalendarApp.OutlookImport, new(SevenA, ImportedAt: At, ImportHash: "h"), At.AddHours(9)));

    // Henter kalenderprogrammet ikke skemaet længere (FetchWatch), kommer knappen igen med en note; har det ikke hentet
    // noget i lang tid, kun noten.
    [Fact]
    public void Not_fetched_any_more()
    {
        var fetched = new Subscription(SevenA, AddedAt: At, FetchedAt: At);
        var missing = new FetchCheck(FetchHealth.Missing, "Outlook", At, null);
        Assert.Equal(new RowButtons("Tilføj til Outlook", null, "Ikke hentet af Outlook siden 3. okt.", "Tilføj igen"),
            RowPresenter.Buttons(CalendarApp.OutlookClassic, fetched, false, Later, Copenhagen, missing));
        Assert.Equal(new RowButtons(null, null, "Ikke hentet af Kalender siden 3. okt.", "Tilføj igen"),
            RowPresenter.Buttons(CalendarApp.AppleCalendar, fetched, false, Later, Copenhagen, new(FetchHealth.Quiet, "Kalender", At, null)));
        // Kun Outlook henter alle kalendere lige tit; andre programmer får kun mærket.
        Assert.Equal(new RowButtons(null, null, "Ikke hentet siden i går kl. 14:00", "Tilføj igen"),
            RowPresenter.Buttons(CalendarApp.Other, fetched, false, At.AddDays(1), Copenhagen, missing with { Program = null }));
        Assert.Equal(new RowButtons(null, null, "Ikke hentet af Kalender siden 3. okt.", "Tilføj igen"),
            RowPresenter.Buttons(CalendarApp.AppleCalendar, fetched, false, Later, Copenhagen, missing with { Program = "Kalender" }));
        Assert.Equal(new RowButtons("Tilføj til Outlook", null, "Ikke hentet af Outlook siden kl. 14:00", "Tilføj igen"),
            RowPresenter.Buttons(CalendarApp.OutlookClassic, fetched, false, At.AddHours(7), Copenhagen, missing));
        Assert.Equal(new RowButtons(null, "✓ Tilføjet", null, "Tilføj igen"),
            RowPresenter.Buttons(CalendarApp.OutlookClassic, fetched, false, Later, Copenhagen, new(FetchHealth.Ok, "Outlook", At, null)));
        // Import: hentning betyder intet.
        Assert.Equal(new RowButtons("Importér…", null, null, "Importér igen…"),
            RowPresenter.Buttons(CalendarApp.OutlookImport, fetched, false, Later, Copenhagen, missing));
    }

    [Fact]
    public void Changes_when_fetch_check_says_so()
    {
        var fetched = new Subscription(SevenA, FetchedAt: At);
        var check = new FetchCheck(FetchHealth.Ok, "Outlook", At, At.AddDays(7));
        Assert.Equal(At.AddDays(7), RowPresenter.ChangesAt(CalendarApp.OutlookClassic, fetched, At, check));
        Assert.Null(RowPresenter.ChangesAt(CalendarApp.OutlookClassic, fetched, At));
        Assert.Null(RowPresenter.ChangesAt(CalendarApp.OutlookImport, fetched, At, check));
    }

    // Ved "Kopiér adresse" ved AulaSync ikke, hvornår adressen sættes ind, så der er ingen frist; men en hentning tæller.
    [Fact]
    public void Copy_address_before_and_after_click()
    {
        Assert.Equal(new RowButtons("Kopiér adresse", null, null, "Tilføj igen"), Buttons(CalendarApp.Other, new(SevenA), Later));
        Assert.Equal(new RowButtons(null, "✓ Kopieret", null, "Tilføj igen"), Buttons(CalendarApp.Other, new(SevenA, AddedAt: At), Later));
        Assert.Equal(new RowButtons(null, "✓ Tilføjet", null, "Tilføj igen"), Buttons(CalendarApp.Other, new(SevenA, AddedAt: At, FetchedAt: At), Later));
    }
}

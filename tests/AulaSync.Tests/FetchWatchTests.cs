using AulaSync.Core;

namespace AulaSync.Tests;

// Henter kalenderprogrammet dine andre skemaer, men ikke dette, er kalenderen nok slettet i programmet (FetchWatch).
public class FetchWatchTests
{
    static readonly DateTimeOffset T0 = new(2026, 10, 5, 6, 0, 0, TimeSpan.Zero);
    static readonly ScheduleRef SevenA = new(ScheduleKind.Group, "3", "7A");
    static readonly ScheduleRef Room = new(ScheduleKind.Resource, "53", "53");

    static Subscription Fetched(ScheduleRef s, string? by, DateTimeOffset at) => new(s, FetchedAt: at, LastFetchedAt: at, FetchedBy: by);

    // Programmet hentede 7A ved T0 og derefter kun de andre skemaer hver halve time fra "from" til "until".
    static FetchWatch OthersFetched(string? program, DateTimeOffset from, DateTimeOffset until)
    {
        var watch = new FetchWatch(T0);
        watch.Record(SevenA.FileName, program, T0);
        for (var t = from; t <= until; t += TimeSpan.FromMinutes(30)) watch.Record(Room.FileName, program, t);
        watch.Record(Room.FileName, program, until);
        return watch;
    }

    [Theory]
    [InlineData("Microsoft Office/16.0 (Windows NT 10.0; Microsoft Outlook 16.0.19127; Pro)", "Outlook")]
    [InlineData("Microsoft Outlook 16.0", "Outlook")]
    [InlineData("macOS/15.6 (24G84) CalendarAgent/1000", "Kalender")]
    [InlineData("iCal/5.0 CFNetwork", "Kalender")]
    [InlineData("Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:128.0) Gecko/20100101 Thunderbird/128.3.0", "Thunderbird")]
    [InlineData("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/141.0 Safari/537.36", null)]
    [InlineData("", null)]
    public void Program_from_user_agent(string agent, string? program) => Assert.Equal(program, FetchWatch.ProgramOf(agent));

    // Outlook hentede 7A ved T0 og har siden, uden ophold, kun hentet de andre skemaer: efter 4 timers hentninger og
    // 6 timer uden 7A er kalenderen nok slettet.
    [Fact]
    public void Missing_when_the_same_program_fetches_the_others_but_not_this()
    {
        var watch = OthersFetched("Outlook", T0.AddMinutes(30), T0.AddHours(5));
        Assert.Equal(new FetchCheck(FetchHealth.Missing, "Outlook", T0, null), watch.Check(Fetched(SevenA, "Outlook", T0), T0.AddHours(6)));
        // Når 7A hentes igen, er alt i orden.
        watch.Record(SevenA.FileName, "Outlook", T0.AddHours(6).AddMinutes(1));
        Assert.Equal(FetchHealth.Ok, watch.Check(Fetched(SevenA, "Outlook", T0), T0.AddHours(6).AddMinutes(2)).Health);
    }

    [Fact]
    public void Not_missing_before_six_hours_and_tells_when_it_changes()
    {
        var watch = OthersFetched("Outlook", T0.AddMinutes(30), T0.AddHours(5));
        Assert.Equal(new FetchCheck(FetchHealth.Ok, "Outlook", T0, T0 + FetchWatch.MissingAfter),
            watch.Check(Fetched(SevenA, "Outlook", T0), T0.AddHours(5).AddMinutes(30)));
    }

    // Under 4 timers hentninger uden ophold er ikke nok; et ophold på mere end 3 timer begynder forfra.
    [Fact]
    public void Not_missing_after_a_short_period()
    {
        var watch = OthersFetched("Outlook", T0.AddHours(10), T0.AddHours(13).AddMinutes(59));
        Assert.Equal(FetchHealth.Ok, watch.Check(Fetched(SevenA, "Outlook", T0), T0.AddHours(14)).Health);
        Assert.Equal(FetchHealth.Ok, OthersFetched("Outlook", T0.AddHours(1), T0.AddHours(3).AddMinutes(30))
            .Check(Fetched(SevenA, "Outlook", T0), T0.AddHours(7)).Health);
    }

    // Har det vist sig, at 7A mangler, huskes det, også når Outlook lukkes og åbnes igen, indtil 7A hentes.
    [Fact]
    public void Missing_is_remembered_until_fetched()
    {
        var watch = OthersFetched("Outlook", T0.AddMinutes(30), T0.AddHours(5));
        Assert.Equal(FetchHealth.Ok, watch.Check(Fetched(SevenA, "Outlook", T0), T0.AddHours(5)).Health);
        watch.Record(Room.FileName, "Outlook", T0.AddHours(20)); // næste morgen
        Assert.Equal(FetchHealth.Missing, watch.Check(Fetched(SevenA, "Outlook", T0), T0.AddHours(20)).Health);
    }

    // Outlook starter og henter de andre skemaer og et øjeblik efter 7A: 7A er hentet i perioden.
    [Fact]
    public void Fetched_in_the_same_period_is_fine()
    {
        var watch = new FetchWatch(T0);
        var start = T0.AddDays(3);
        watch.Record(Room.FileName, "Outlook", start);
        watch.Record(SevenA.FileName, "Outlook", start.AddSeconds(1));
        watch.Record(Room.FileName, "Outlook", start.AddHours(3));
        var s = Fetched(SevenA, "Outlook", T0); // gemt tidligere; FetchWatch kender den nyeste
        Assert.Equal(start.AddSeconds(1), watch.LastFetched(s));
        Assert.Equal(FetchHealth.Ok, watch.Check(s, start.AddHours(8)).Health);
    }

    // Et andet program (Thunderbird) siger intet om et skema, Outlook henter.
    [Fact]
    public void Another_program_says_nothing()
    {
        var watch = OthersFetched("Thunderbird", T0.AddMinutes(30), T0.AddHours(8));
        Assert.Equal(FetchHealth.Ok, watch.Check(Fetched(SevenA, "Outlook", T0), T0.AddHours(9)).Health);
    }

    // Apple Kalender kan hente én gang om ugen pr. kalender: først efter 8 dage.
    [Fact]
    public void Calendar_on_the_mac_gets_a_week()
    {
        var watch = OthersFetched("Kalender", T0.AddMinutes(30), T0.AddDays(8));
        Assert.Equal(FetchHealth.Ok, watch.Check(Fetched(SevenA, "Kalender", T0), T0.AddDays(8).AddMinutes(-1)).Health);
        Assert.Equal(FetchHealth.Missing, watch.Check(Fetched(SevenA, "Kalender", T0), T0.AddDays(8)).Health);
    }

    // Ingen hentninger i 3 dage (fx har Outlook været lukket): en note, men først når AulaSync har kørt i 2 timer.
    [Fact]
    public void Quiet_after_a_week_without_fetches()
    {
        var now = T0.AddDays(10);
        var justStarted = new FetchWatch(now.AddHours(-1));
        var check = justStarted.Check(Fetched(SevenA, "Outlook", T0), now);
        Assert.Equal(new FetchCheck(FetchHealth.Ok, "Outlook", T0, now.AddHours(1)), check); // efter 2 timer
        var running = new FetchWatch(now.AddHours(-2));
        Assert.Equal(new FetchCheck(FetchHealth.Quiet, "Outlook", T0, null), running.Check(Fetched(SevenA, "Outlook", T0), now));
        Assert.Equal(T0 + FetchWatch.QuietAfter, new FetchWatch(T0).Check(Fetched(SevenA, "Outlook", T0), T0.AddDays(1)).ChangesAt);
    }

    // Et ur, der har været stillet frem, giver ikke en hentning i fremtiden.
    [Fact]
    public void Future_fetch_counts_as_now()
    {
        var check = new FetchWatch(T0).Check(Fetched(SevenA, "Outlook", T0.AddDays(60)), T0);
        Assert.Equal((FetchHealth.Ok, T0, T0 + FetchWatch.QuietAfter), (check.Health, check.LastFetched, check.ChangesAt));
    }

    // Uden hentning (eller efter "Tilføj igen") har FetchWatch intet at sige; rækken venter som før.
    [Fact]
    public void Nothing_to_say_before_the_first_fetch()
    {
        var watch = new FetchWatch(T0);
        Assert.Equal(FetchCheck.None, watch.Check(new Subscription(SevenA), T0.AddDays(10)));
        Assert.Equal(FetchCheck.None, watch.Check(new Subscription(SevenA, AddedAt: T0.AddDays(1), FetchedAt: T0), T0.AddDays(10)));
    }
}

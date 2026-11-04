using System.Text;
using AulaSync.Core;
using static AulaSync.Tests.TestData;

namespace AulaSync.Tests;

public class IcsWriterTests
{
    static readonly DateTimeOffset Now = new(2026, 10, 6, 12, 0, 0, TimeSpan.Zero);
    static readonly ScheduleRef Anna = new(ScheduleKind.Employee, "1000001", "Anna Eksempel", "AE");

    static string[] Lines(string ics) => ics.Split("\r\n");

    [Fact]
    public void Empty_calendar_is_valid()
    {
        var ics = IcsWriter.Write(Anna, [], Now);
        Assert.StartsWith("BEGIN:VCALENDAR\r\nVERSION:2.0\r\n", ics);
        Assert.EndsWith("END:VCALENDAR\r\n", ics);
        Assert.Contains("PRODID:-//AulaSync//Skema 3.0//DA\r\n", ics);
        Assert.Contains("CALSCALE:GREGORIAN\r\n", ics);
        Assert.Contains("METHOD:PUBLISH\r\n", ics);
        Assert.Contains("X-WR-CALNAME:AE Anna Eksempel\r\n", ics);
        Assert.Contains("REFRESH-INTERVAL;VALUE=DURATION:PT1H\r\n", ics);
        Assert.Contains("X-PUBLISHED-TTL:PT1H\r\n", ics);
        Assert.DoesNotContain("BEGIN:VEVENT", ics);
    }

    [Fact]
    public void Event_has_required_fields_in_utc()
    {
        var ics = IcsWriter.Write(Anna, [Ev("DAN", "53", ["7A"], [("Anna Eksempel", "AE")], id: "5001")], Now);
        Assert.Contains("BEGIN:VEVENT\r\n", ics);
        Assert.Contains("UID:aula-5001-20260413T060000Z-medarbejder-1000001@aulasync\r\n", ics);
        Assert.Contains("DTSTAMP:20261006T120000Z\r\n", ics);
        Assert.Contains("DTSTART:20260413T060000Z\r\n", ics);
        Assert.Contains("DTEND:20260413T064500Z\r\n", ics);
        Assert.Contains("SUMMARY:DAN | 53 | 7A\r\n", ics);
        Assert.Contains("LOCATION:53\r\n", ics);
        Assert.Contains(@"DESCRIPTION:Anna Eksempel (AE)\nKlasse: 7A\nLokale: 53" + "\r\n", ics);
        Assert.Contains("TRANSP:TRANSPARENT\r\n", ics);
        Assert.DoesNotContain("VALARM", ics);
    }

    [Fact]
    public void Uid_differs_per_calendar_for_same_event()
    {
        var e = Ev(id: "5001");
        var a = IcsWriter.Write(Anna, [e], Now);
        var b = IcsWriter.Write(new ScheduleRef(ScheduleKind.Group, "88", "7A"), [e], Now);
        Assert.Contains("UID:aula-5001-20260413T060000Z-klasse-88@aulasync", b);
        Assert.DoesNotContain("UID:aula-5001-20260413T060000Z-klasse-88@aulasync", a);
    }

    // Gentagne begivenheder deler id i Aula (webappen bruger id + starttid som nøgle); hver forekomst skal have sit eget UID.
    [Fact]
    public void Recurring_occurrences_get_distinct_uids()
    {
        var ics = IcsWriter.Write(Anna,
        [
            Ev(id: "900", start: "2026-04-13T14:00:00+02:00", end: "2026-04-13T15:00:00+02:00"),
            Ev(id: "900", start: "2026-04-20T14:00:00+02:00", end: "2026-04-20T15:00:00+02:00"),
        ], Now);
        Assert.Contains("UID:aula-900-20260413T120000Z-medarbejder-1000001@aulasync\r\n", ics);
        Assert.Contains("UID:aula-900-20260420T120000Z-medarbejder-1000001@aulasync\r\n", ics);
    }

    [Fact]
    public void Winter_time_is_converted_correctly()
    {
        var ics = IcsWriter.Write(Anna, [Ev(start: "2026-01-05T08:00:00+01:00", end: "2026-01-05T08:45:00+01:00")], Now);
        Assert.Contains("DTSTART:20260105T070000Z\r\n", ics);
    }

    [Fact]
    public void End_not_after_start_omits_dtend()
    {
        var ics = IcsWriter.Write(Anna, [Ev(start: "2026-04-13T08:00:00+02:00", end: "2026-04-13T08:00:00+02:00")], Now);
        Assert.Contains("DTSTART:20260413T060000Z\r\n", ics);
        Assert.DoesNotContain("DTEND", ics);
    }

    [Fact]
    public void Special_characters_are_escaped()
    {
        var ics = IcsWriter.Write(Anna, [Ev("Test; med, special\\tegn", "Rum; 1,2")], Now);
        Assert.Contains(@"SUMMARY:Test\; med\, special\\tegn | Rum\; 1\,2", ics);
        Assert.Contains(@"LOCATION:Rum\; 1\,2", ics);
    }

    [Fact]
    public void Calendar_name_is_escaped()
    {
        var ics = IcsWriter.Write(new ScheduleRef(ScheduleKind.Resource, "4", "Sal 1, 2; 3"), [], Now);
        Assert.Contains(@"X-WR-CALNAME:Sal 1\, 2\; 3", ics);
    }

    [Fact]
    public void Only_crlf_line_endings_and_no_line_over_75_bytes()
    {
        var many = Enumerable.Range(0, 20).Select(i => ($"Lærer Ærø Østergård nummer {i}", $"L{i}")).ToArray();
        var ics = IcsWriter.Write(Anna, [Ev("Særlig dansktime med æøå og en meget lang titel der fortsætter", "Lokale ÆØÅ", ["7A", "7B", "7C"], many)], Now);
        Assert.DoesNotContain("\n", ics.Replace("\r\n", ""));
        Assert.All(Lines(ics), l => Assert.True(Encoding.UTF8.GetByteCount(l) <= 75, l));
    }

    [Fact]
    public void Events_are_ordered_by_start_then_id()
    {
        var ics = IcsWriter.Write(Anna,
        [
            Ev(id: "3", start: "2026-04-13T10:00:00+02:00", end: "2026-04-13T10:45:00+02:00"),
            Ev(id: "2", start: "2026-04-13T08:00:00+02:00", end: "2026-04-13T08:45:00+02:00"),
            Ev(id: "1", start: "2026-04-13T08:00:00+02:00", end: "2026-04-13T08:45:00+02:00"),
        ], Now);
        var uids = Lines(ics).Where(l => l.StartsWith("UID:")).ToList();
        Assert.Equal(["UID:aula-1-20260413T060000Z-medarbejder-1000001@aulasync", "UID:aula-2-20260413T060000Z-medarbejder-1000001@aulasync", "UID:aula-3-20260413T080000Z-medarbejder-1000001@aulasync"], uids);
    }

    [Fact]
    public void Every_event_is_closed()
    {
        var ics = IcsWriter.Write(Anna, [Ev(id: "1"), Ev(id: "2"), Ev(id: "3")], Now);
        Assert.Equal(3, Lines(ics).Count(l => l == "BEGIN:VEVENT"));
        Assert.Equal(3, Lines(ics).Count(l => l == "END:VEVENT"));
    }
    [Theory]
    [InlineData(30, "PT30M")]
    [InlineData(60, "PT1H")]
    public void Refresh_follows_options(int minutes, string expected)
    {
        var ics = IcsWriter.Write(Anna, [], Now, new IcsOptions(TimeSpan.FromMinutes(minutes)));
        Assert.Contains($"REFRESH-INTERVAL;VALUE=DURATION:{expected}\r\n", ics);
        Assert.Contains($"X-PUBLISHED-TTL:{expected}\r\n", ics);
    }

    // Statusbegivenheden for mandagen med datoen (yyyyMMdd).
    static string StatusEvent(DateTimeOffset now, string monday, TimeOnly? time = null)
    {
        var ics = IcsWriter.Write(Anna, [Ev(id: "5001")], now, new IcsOptions(TimeSpan.FromHours(1), Copenhagen, time));
        var start = ics.IndexOf($"BEGIN:VEVENT\r\nUID:aulasync-status-medarbejder-1000001-{monday}@aulasync", StringComparison.Ordinal);
        Assert.True(start >= 0, $"Ingen statusbegivenhed mandag {monday}");
        return ics[start..(ics.IndexOf("END:VEVENT\r\n", start, StringComparison.Ordinal) + 12)];
    }

    // Slået til i Indstillinger: en privat begivenhed mandag kl. 5.45-6.00 viser, hvornår skemaet blev hentet, med samme
    // tekst i titel og beskrivelse.
    [Fact]
    public void Status_event_is_private_on_monday_morning()
    {
        var status = StatusEvent(new DateTimeOffset(2026, 10, 7, 12, 32, 0, TimeSpan.Zero), "20261005"); // onsdag 14:32 dansk tid
        Assert.Equal(string.Join("\r\n",
            "BEGIN:VEVENT",
            "UID:aulasync-status-medarbejder-1000001-20261005@aulasync",
            "DTSTAMP:20261007T123200Z",
            "DTSTART:20261005T034500Z",
            "DTEND:20261005T040000Z",
            "SUMMARY:Opd. 071026@14:32",
            "DESCRIPTION:Opd. 071026@14:32",
            "CLASS:PRIVATE",
            "TRANSP:TRANSPARENT",
            "END:VEVENT") + "\r\n", Unfold(status));
    }

    // Teksten er hentetidspunktet i dansk tid; vinter- og sommertid regnes om til UTC for hver mandag for sig.
    [Theory]
    [InlineData("2026-10-05T05:00:00+02:00", "20261005", "20261005T034500Z", "Opd. 051026@05:00")]
    [InlineData("2026-10-09T22:30:00Z", "20261012", "20261012T034500Z", "Opd. 101026@00:30")]
    [InlineData("2026-10-10T08:00:00+02:00", "20261012", "20261012T034500Z", "Opd. 101026@08:00")]
    [InlineData("2026-10-25T12:00:00+01:00", "20261026", "20261026T044500Z", "Opd. 251026@12:00")]
    [InlineData("2026-11-03T09:15:00+01:00", "20261102", "20261102T044500Z", "Opd. 031126@09:15")]
    public void Status_event_text_and_time_zone(string now, string monday, string start, string text)
    {
        var status = Unfold(StatusEvent(DateTimeOffset.Parse(now), monday));
        Assert.Contains($"DTSTART:{start}\r\n", status);
        Assert.Contains($"SUMMARY:{text}\r\n", status);
        Assert.Contains($"DESCRIPTION:{text}\r\n", status);
    }

    // Hver mandag i skemaets periode (90 dage tilbage og frem), så den ses i hver uge, med datoen i UID'en og samme tekst.
    [Fact]
    public void Status_event_every_monday_in_the_period()
    {
        var now = new DateTimeOffset(2026, 10, 7, 12, 32, 0, TimeSpan.Zero); // onsdag 14:32 dansk tid
        var ics = Unfold(IcsWriter.Write(Anna, [Ev(id: "5001")], now, new IcsOptions(TimeSpan.FromHours(1), Copenhagen)));
        var starts = System.Text.RegularExpressions.Regex.Matches(ics,
                @"UID:aulasync-status-medarbejder-1000001-(\d{8})@aulasync\r\nDTSTAMP:\S+\r\nDTSTART:(\S+)\r\n")
            .Select(m => (Day: m.Groups[1].Value, Start: m.Groups[2].Value)).ToList();
        // 9. juli 2026 (90 dage før) er en torsdag: første mandag er 13. juli; 5. januar 2027 (90 dage efter) er en tirsdag:
        // sidste mandag er 4. januar.
        Assert.Equal(26, starts.Count);
        Assert.Equal(("20260713", "20260713T034500Z"), starts[0]);
        Assert.Contains(("20261019", "20261019T034500Z"), starts); // sommertid
        Assert.Contains(("20261026", "20261026T044500Z"), starts); // vintertid fra 25. oktober
        Assert.Equal(("20270104", "20270104T044500Z"), starts[^1]);
        Assert.Equal(26, System.Text.RegularExpressions.Regex.Matches(ics, "SUMMARY:Opd. 071026@14:32\r\n").Count);
        Assert.Equal(26, System.Text.RegularExpressions.Regex.Matches(ics, "CLASS:PRIVATE\r\n").Count);
        // Lektionerne tælles og hashes som uden statusbegivenheder (Ændret siden import).
        var plain = IcsWriter.Write(Anna, [Ev(id: "5001")], now, new IcsOptions(TimeSpan.FromHours(1)));
        Assert.Equal(IcsInspect.Read(plain).ContentHash, IcsInspect.Read(ics).ContentHash);
        Assert.Equal(1, IcsInspect.Read(ics).Events);
    }

    // Klokkeslættet fra Indstillinger, i dansk tid på begge sider af skiftet til vintertid; et kvarter langt.
    [Theory]
    [InlineData(7, 30, "20261005", "20261005T053000Z", "20261005T054500Z")]
    [InlineData(7, 30, "20261102", "20261102T063000Z", "20261102T064500Z")]
    [InlineData(23, 55, "20261005", "20261005T215500Z", "20261005T221000Z")]
    public void Status_event_at_chosen_time(int hour, int minute, string monday, string start, string end)
    {
        var status = Unfold(StatusEvent(new DateTimeOffset(2026, 10, 7, 12, 32, 0, TimeSpan.Zero), monday, new TimeOnly(hour, minute)));
        Assert.Contains($"DTSTART:{start}\r\n", status);
        Assert.Contains($"DTEND:{end}\r\n", status);
    }

    // Heldagsbegivenhed: hele mandagen (dato uden klokkeslæt), ellers som aftalen.
    [Fact]
    public void Status_event_as_all_day()
    {
        var now = new DateTimeOffset(2026, 10, 7, 12, 32, 0, TimeSpan.Zero);
        var ics = Unfold(IcsWriter.Write(Anna, [Ev(id: "5001")], now,
            new IcsOptions(TimeSpan.FromHours(1), Copenhagen, StatusStyle: StatusStyle.AllDay)));
        var start = ics.IndexOf("UID:aulasync-status-medarbejder-1000001-20261026@aulasync", StringComparison.Ordinal);
        Assert.True(start >= 0);
        Assert.Equal(string.Join("\r\n",
            "UID:aulasync-status-medarbejder-1000001-20261026@aulasync",
            "DTSTAMP:20261007T123200Z",
            "DTSTART;VALUE=DATE:20261026",
            "DTEND;VALUE=DATE:20261027",
            "SUMMARY:Opd. 071026@14:32",
            "DESCRIPTION:Opd. 071026@14:32",
            "CLASS:PRIVATE",
            "TRANSP:TRANSPARENT",
            "END:VEVENT") + "\r\n", ics[start..(ics.IndexOf("END:VEVENT\r\n", start, StringComparison.Ordinal) + 12)]);
        Assert.Equal(26, System.Text.RegularExpressions.Regex.Matches(ics, "DTSTART;VALUE=DATE:").Count);
    }

    // Datoerne er gregorianske, også når Windows bruger fx thailandsk eller persisk format.
    [Theory]
    [InlineData("th-TH")]
    [InlineData("fa-IR")]
    public void Status_dates_ignore_the_regional_calendar(string culture)
    {
        var old = System.Globalization.CultureInfo.CurrentCulture;
        try
        {
            System.Globalization.CultureInfo.CurrentCulture = new System.Globalization.CultureInfo(culture);
            var now = new DateTimeOffset(2026, 10, 7, 12, 32, 0, TimeSpan.Zero);
            var allDay = IcsWriter.Write(Anna, [], now, new IcsOptions(TimeSpan.FromHours(1), Copenhagen, StatusStyle: StatusStyle.AllDay));
            Assert.Contains("UID:aulasync-status-medarbejder-1000001-20261026@aulasync\r\nDTSTAMP:20261007T123200Z\r\nDTSTART;VALUE=DATE:20261026\r\nDTEND;VALUE=DATE:20261027\r\n", allDay);
            Assert.Contains("SUMMARY:Opd. 071026@14:32\r\n", allDay);
            var timed = IcsWriter.Write(Anna, [], now, new IcsOptions(TimeSpan.FromHours(1), Copenhagen));
            Assert.Contains("UID:aulasync-status-medarbejder-1000001-20261026@aulasync\r\n", timed);
        }
        finally { System.Globalization.CultureInfo.CurrentCulture = old; }
    }

    // I hver lektions beskrivelse: ingen statusbegivenheder; tidspunktet står sidst efter en tom linje, også i en lektion
    // uden beskrivelse. Antal og hash er de samme som uden (Ændret siden import), uanset hvordan tidspunktet vises.
    [Fact]
    public void Status_in_every_lesson_description()
    {
        var now = new DateTimeOffset(2026, 10, 7, 12, 32, 0, TimeSpan.Zero);
        AulaEvent[] lessons = [Ev(id: "5001"), Ev(id: "5002", title: "Pause", location: "", groups: [], teachers: [], start: "2026-04-13T10:00:00+02:00", end: "2026-04-13T10:30:00+02:00")];
        IcsOptions Style(StatusStyle s) => new(TimeSpan.FromHours(1), Copenhagen, StatusStyle: s);
        var ics = Unfold(IcsWriter.Write(Anna, lessons, now, Style(StatusStyle.InDescription)));
        Assert.DoesNotContain("aulasync-status", ics);
        Assert.Matches(@"DESCRIPTION:.+\\n\\nOpd\. 071026@14:32\r\n", ics);
        Assert.Contains("DESCRIPTION:Opd. 071026@14:32\r\n", ics); // pausen uden beskrivelse
        var plain = IcsInspect.Read(IcsWriter.Write(Anna, lessons, now, new IcsOptions(TimeSpan.FromHours(1))));
        Assert.Equal(2, plain.Events);
        foreach (var style in Enum.GetValues<StatusStyle>())
        {
            var summary = IcsInspect.Read(IcsWriter.Write(Anna, lessons, now.AddHours(3), Style(style)));
            Assert.Equal((plain.Events, plain.ContentHash), (summary.Events, summary.ContentHash));
        }
    }

    // Hvornår skemaet blev hentet, står i filens hoved, så det kan læses, også når opdateringstiden er skrevet om.
    [Fact]
    public void Fetch_time_is_in_the_header()
    {
        var ics = IcsWriter.Write(Anna, [Ev(id: "5001")], Now);
        Assert.Contains("X-PUBLISHED-TTL:PT1H\r\nX-AULASYNC-FETCHED:20261006T120000Z\r\n", ics);
        Assert.Equal(Now, IcsInspect.Read(ics).Fetched);
    }

    static readonly DateTimeOffset Fetched = new(2026, 10, 7, 12, 32, 0, TimeSpan.Zero); // onsdag 14:32 dansk tid
    static readonly AulaEvent[] Lessons = [Ev(id: "5001"), Ev(id: "5002", title: "Pause", location: "", groups: [], teachers: [],
        start: "2026-04-13T10:00:00+02:00", end: "2026-04-13T10:30:00+02:00")];

    static string Stampless(string ics) =>
        string.Join("\r\n", Unfold(ics).Split("\r\n").Where(l => !l.StartsWith("DTSTAMP:", StringComparison.Ordinal)));

    // Logget ud af Aula (Rewrite): lektionerne bliver, og opdateringstiden fra hentningen får "⚠ Log ind i AulaSync · " foran,
    // på den måde, der er valgt. Antal, hash og hentetidspunkt er de samme (Ændret siden import), DTSTAMP er nu, og logget ind
    // igen er filen som før, bortset fra DTSTAMP.
    [Theory]
    [InlineData(StatusStyle.Timed, 52)]
    [InlineData(StatusStyle.AllDay, 52)]
    [InlineData(StatusStyle.InDescription, 2)]
    public void Rewrite_shows_logged_out_and_back(StatusStyle style, int warnings)
    {
        var later = Fetched.AddDays(2);
        var options = new IcsOptions(TimeSpan.FromHours(1), Copenhagen, StatusStyle: style);
        var ics = IcsWriter.Write(Anna, Lessons, Fetched, options);

        var warned = IcsWriter.Rewrite(ics, Anna, later, options, loggedOut: true, fetchedFallback: later)!;
        Assert.Equal(warnings, System.Text.RegularExpressions.Regex.Matches(Unfold(warned), "⚠ Log ind i AulaSync · Opd\\. 071026@14:32(\r\n)").Count);
        Assert.DoesNotContain(":Opd. ", Unfold(warned));
        Assert.DoesNotContain("\\n\\nOpd. ", Unfold(warned));
        Assert.Contains("DTSTAMP:20261009T123200Z\r\n", warned);
        Assert.DoesNotContain("DTSTAMP:20261007T123200Z", warned);
        var (before, after) = (IcsInspect.Read(ics), IcsInspect.Read(warned));
        Assert.Equal((before.Events, before.ContentHash, Fetched), (after.Events, after.ContentHash, after.Fetched));
        Assert.Null(IcsWriter.Rewrite(warned, Anna, later.AddHours(1), options, loggedOut: true, fetchedFallback: later));

        var back = IcsWriter.Rewrite(warned, Anna, later.AddHours(1), options, loggedOut: false, fetchedFallback: later)!;
        Assert.Equal(Stampless(ics), Stampless(back));
    }

    // En fil fra før 3.4 har ikke hentetidspunktet i hovedet: det er lektionernes DTSTAMP, og hovedet får det.
    [Fact]
    public void Rewrite_of_a_file_from_before_3_4_uses_dtstamp()
    {
        var options = new IcsOptions(TimeSpan.FromHours(1), Copenhagen);
        var old = IcsWriter.Write(Anna, Lessons, Fetched, options).Replace("X-AULASYNC-FETCHED:20261007T123200Z\r\n", "");
        Assert.Equal(Fetched, IcsInspect.Read(old).Fetched);
        var warned = IcsWriter.Rewrite(old, Anna, Fetched.AddDays(1), options, loggedOut: true, fetchedFallback: Fetched.AddDays(1))!;
        Assert.Contains("X-PUBLISHED-TTL:PT1H\r\nX-AULASYNC-FETCHED:20261007T123200Z\r\nBEGIN:VEVENT\r\n", warned);
        Assert.Contains("SUMMARY:⚠ Log ind i AulaSync · Opd. 071026@14:32\r\n", Unfold(warned));
        Assert.Equal(Fetched, IcsInspect.Read(warned).Fetched);
    }

    // Uden lektioner og uden hentetidspunkt i filen bruges fetchedFallback (fx hvornår filen blev skrevet).
    [Fact]
    public void Rewrite_of_an_empty_old_file_uses_the_fallback()
    {
        var options = new IcsOptions(TimeSpan.FromHours(1), Copenhagen);
        var old = IcsWriter.Write(Anna, [], Fetched).Replace("X-AULASYNC-FETCHED:20261007T123200Z\r\n", "");
        var warned = IcsWriter.Rewrite(old, Anna, Fetched.AddDays(1), options, loggedOut: true, fetchedFallback: Fetched.AddHours(-1))!;
        Assert.Contains("SUMMARY:⚠ Log ind i AulaSync · Opd. 071026@13:32\r\n", Unfold(warned));
    }

    // Er Vis opdateringstid i kalenderen slået fra, er der intet at vise: opdateringstiden forsvinder, og står der ingen,
    // ændres intet.
    [Theory]
    [InlineData(StatusStyle.Timed)]
    [InlineData(StatusStyle.InDescription)]
    public void Rewrite_without_status_removes_it(StatusStyle style)
    {
        var off = new IcsOptions(TimeSpan.FromHours(1));
        var ics = IcsWriter.Write(Anna, Lessons, Fetched, new IcsOptions(TimeSpan.FromHours(1), Copenhagen, StatusStyle: style));
        var removed = IcsWriter.Rewrite(ics, Anna, Fetched.AddDays(1), off, loggedOut: true, fetchedFallback: Fetched)!;
        Assert.Equal(Stampless(IcsWriter.Write(Anna, Lessons, Fetched, off)), Stampless(removed));
        Assert.Null(IcsWriter.Rewrite(removed, Anna, Fetched.AddDays(2), off, loggedOut: true, fetchedFallback: Fetched));
    }

    [Theory]
    [InlineData("")]
    [InlineData("ikke en kalender")]
    [InlineData("BEGIN:VCALENDAR\r\nVERSION:2.0\r\n")]
    public void Rewrite_leaves_other_files_alone(string text) =>
        Assert.Null(IcsWriter.Rewrite(text, Anna, Fetched, new IcsOptions(TimeSpan.FromHours(1), Copenhagen), true, Fetched));

    [Fact]
    public void No_status_event_by_default() =>
        Assert.DoesNotContain("aulasync-status", IcsWriter.Write(Anna, [Ev()], Now));

    static string Unfold(string ics) => ics.Replace("\r\n ", "");
}

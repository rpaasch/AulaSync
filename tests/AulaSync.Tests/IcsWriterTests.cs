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

    static string StatusEvent(DateTimeOffset now)
    {
        var ics = IcsWriter.Write(Anna, [Ev(id: "5001")], now, new IcsOptions(TimeSpan.FromHours(1), Copenhagen));
        var start = ics.IndexOf("BEGIN:VEVENT\r\nUID:aulasync-status-", StringComparison.Ordinal);
        Assert.True(start >= 0, "Ingen statusbegivenhed");
        return ics[start..(ics.IndexOf("END:VEVENT\r\n", start, StringComparison.Ordinal) + 12)];
    }

    // Slået til i Indstillinger: en privat begivenhed mandag kl. 5.45-6.00 viser, hvornår skemaet blev hentet.
    [Fact]
    public void Status_event_is_private_on_monday_morning()
    {
        var status = StatusEvent(new DateTimeOffset(2026, 10, 7, 12, 32, 0, TimeSpan.Zero)); // onsdag 14:32 dansk tid
        Assert.Equal(string.Join("\r\n",
            "BEGIN:VEVENT",
            "UID:aulasync-status-medarbejder-1000001@aulasync",
            "DTSTAMP:20261007T123200Z",
            "DTSTART:20261005T034500Z",
            "DTEND:20261005T040000Z",
            "SUMMARY:AulaSync opdateret ons. 7. okt. 14:32",
            @"DESCRIPTION:AulaSync hentede skemaet fra Aula onsdag 7. oktober 2026 kl. 14:32.\nDenne private aftale kan slås fra i AulaSync under Indstillinger.",
            "CLASS:PRIVATE",
            "TRANSP:TRANSPARENT",
            "END:VEVENT") + "\r\n", Unfold(status));
    }

    // Lørdag og søndag ligger den i den kommende uge; vintertid regnes om til UTC.
    [Theory]
    [InlineData("2026-10-05T05:00:00+02:00", "20261005T034500Z", "man. 5. okt. 05:00")]
    [InlineData("2026-10-09T23:59:00+02:00", "20261005T034500Z", "fre. 9. okt. 23:59")]
    [InlineData("2026-10-10T08:00:00+02:00", "20261012T034500Z", "lør. 10. okt. 08:00")]
    [InlineData("2026-10-25T12:00:00+01:00", "20261026T044500Z", "søn. 25. okt. 12:00")]
    [InlineData("2026-11-03T09:15:00+01:00", "20261102T044500Z", "tirs. 3. nov. 09:15")]
    [InlineData("2026-11-05T09:15:00+01:00", "20261102T044500Z", "tors. 5. nov. 09:15")]
    public void Status_event_week_and_time_zone(string now, string start, string stamp)
    {
        var status = Unfold(StatusEvent(DateTimeOffset.Parse(now)));
        Assert.Contains($"DTSTART:{start}\r\n", status);
        Assert.Contains($"SUMMARY:AulaSync opdateret {stamp}\r\n", status);
    }

    [Fact]
    public void No_status_event_by_default() =>
        Assert.DoesNotContain("aulasync-status", IcsWriter.Write(Anna, [Ev()], Now));

    static string Unfold(string ics) => ics.Replace("\r\n ", "");
}

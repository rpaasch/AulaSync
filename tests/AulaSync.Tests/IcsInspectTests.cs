using AulaSync.Core;
using static AulaSync.Tests.TestData;

namespace AulaSync.Tests;

public class IcsInspectTests
{
    static readonly ScheduleRef SevenA = new(ScheduleKind.Group, "88231", "7A");
    static readonly DateTimeOffset Now = new(2026, 10, 6, 12, 0, 0, TimeSpan.Zero);
    static AulaEvent Lesson(string id, string start = "2026-04-13T08:00:00+02:00", string end = "2026-04-13T08:45:00+02:00") => Ev(id: id, start: start, end: end);

    [Fact]
    public void Counts_events() =>
        Assert.Equal(3, IcsInspect.Read(IcsWriter.Write(SevenA, [Lesson("1"), Lesson("2"), Lesson("3")], Now)).Events);

    [Fact]
    public void Empty_calendar_has_zero_events() =>
        Assert.Equal(0, IcsInspect.Read(IcsWriter.Write(SevenA, [], Now)).Events);

    [Fact]
    public void Hash_ignores_dtstamp()
    {
        var a = IcsInspect.Read(IcsWriter.Write(SevenA, [Lesson("1")], Now)).ContentHash;
        var b = IcsInspect.Read(IcsWriter.Write(SevenA, [Lesson("1")], Now.AddHours(6))).ContentHash;
        Assert.Equal(a, b);
    }

    [Fact]
    public void Hash_ignores_order()
    {
        var a = IcsInspect.Read(IcsWriter.Write(SevenA, [Lesson("1"), Lesson("2")], Now)).ContentHash;
        var b = IcsInspect.Read(IcsWriter.Write(SevenA, [Lesson("2"), Lesson("1")], Now)).ContentHash;
        Assert.Equal(a, b);
    }

    [Fact]
    public void Hash_changes_when_lesson_moves_is_added_or_removed()
    {
        var baseline = IcsInspect.Read(IcsWriter.Write(SevenA, [Lesson("1"), Lesson("2")], Now)).ContentHash;
        Assert.NotEqual(baseline, IcsInspect.Read(IcsWriter.Write(SevenA, [Lesson("1"), Lesson("2", "2026-04-13T10:00:00+02:00", "2026-04-13T10:45:00+02:00")], Now)).ContentHash);
        Assert.NotEqual(baseline, IcsInspect.Read(IcsWriter.Write(SevenA, [Lesson("1"), Lesson("2"), Lesson("3")], Now)).ContentHash);
        Assert.NotEqual(baseline, IcsInspect.Read(IcsWriter.Write(SevenA, [Lesson("1")], Now)).ContentHash);
    }

    [Fact]
    public void Missing_file_is_null()
    {
        using var dir = new TempDir();
        Assert.Null(IcsInspect.ReadFile(dir.File("findes-ikke.ics")));
    }

    // Kun begivenheder, der starter i tidsrummet (begge grænser med), tæller.
    [Fact]
    public void Range_counts_only_events_starting_inside_it()
    {
        var ics = IcsWriter.Write(SevenA, [
            Lesson("1", "2026-04-13T08:00:00+02:00", "2026-04-13T08:45:00+02:00"),
            Lesson("2", "2026-04-14T08:00:00+02:00", "2026-04-14T08:45:00+02:00"),
            Lesson("3", "2026-04-15T08:00:00+02:00", "2026-04-15T08:45:00+02:00")], Now);
        var from = DateTimeOffset.Parse("2026-04-14T08:00:00+02:00");
        Assert.Equal(1, IcsInspect.Read(ics, from, from).Events);

        var onward = IcsInspect.Read(ics, from);
        Assert.Equal(2, onward.Events);
        Assert.Equal(onward.ContentHash, IcsInspect.Read(ics, from, DateTimeOffset.Parse("2026-04-15T08:00:00+02:00")).ContentHash);
        Assert.NotEqual(IcsInspect.Read(ics).ContentHash, onward.ContentHash);
        Assert.Equal(0, IcsInspect.Read(ics, DateTimeOffset.Parse("2026-05-01T00:00:00Z")).Events);
    }
}

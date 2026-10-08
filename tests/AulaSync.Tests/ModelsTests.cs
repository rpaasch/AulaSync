using AulaSync.Core;

namespace AulaSync.Tests;

public class ModelsTests
{
    [Theory]
    [InlineData(ScheduleKind.Employee, "1000001", "medarbejder-1000001.ics")]
    [InlineData(ScheduleKind.Group, "88231", "klasse-88231.ics")]
    [InlineData(ScheduleKind.Resource, "412", "lokale-412.ics")]
    public void FileName_uses_kind_and_id(ScheduleKind kind, string id, string expected) =>
        Assert.Equal(expected, new ScheduleRef(kind, id, "X").FileName);

    [Fact]
    public void Employee_calendar_name_has_initials_first() =>
        Assert.Equal("AE Anna Eksempel", new ScheduleRef(ScheduleKind.Employee, "1", "Anna Eksempel", "AE").CalendarName);

    [Fact]
    public void Group_calendar_name_is_name() =>
        Assert.Equal("7A", new ScheduleRef(ScheduleKind.Group, "1", "7A").CalendarName);

    [Fact]
    public void Employee_without_initials_uses_name() =>
        Assert.Equal("Anna Eksempel", new ScheduleRef(ScheduleKind.Employee, "1", "Anna Eksempel").CalendarName);

    [Theory]
    [InlineData("")]
    [InlineData("../x")]
    [InlineData("a/b")]
    [InlineData("1 2")]
    [InlineData("æ")]
    public void Invalid_id_is_rejected(string id)
    {
        Assert.False(ScheduleRef.IsValidId(id));
        Assert.Throws<ArgumentException>(() => new ScheduleRef(ScheduleKind.Group, id, "X"));
    }

    [Fact]
    public void Key_ignores_name() =>
        Assert.Equal(new ScheduleRef(ScheduleKind.Group, "5", "7A").Key, new ScheduleRef(ScheduleKind.Group, "5", "7.A").Key);

    [Fact]
    public void AppPaths_layout()
    {
        var p = new AppPaths("/root");
        Assert.Equal(Path.Combine("/root", "kalendere"), p.Calendars);
        Assert.Equal(Path.Combine("/root", "abonnementer.json"), p.Subscriptions);
        Assert.Equal(Path.Combine("/root", "config.json"), p.Config);
        Assert.Equal(Path.Combine("/root", "aulasync.log"), p.Log);
    }
}

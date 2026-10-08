using AulaSync.Core;

namespace AulaSync.App.Tests;

public class BadgeTests
{
    [Theory]
    [InlineData(ScheduleKind.Employee, "Anna Eksempel", "AE", "AE")]
    [InlineData(ScheduleKind.Employee, "Bo Testesen", "", "BT")]
    [InlineData(ScheduleKind.Group, "7A", "", "7A")]
    [InlineData(ScheduleKind.Group, "10A", "", "10A")]
    [InlineData(ScheduleKind.Group, "7 årgang", "", "7år")]
    [InlineData(ScheduleKind.Resource, "Lokale 53", "", "53")]
    [InlineData(ScheduleKind.Resource, "Gymnastiksal", "", "Gym")]
    public void Text(ScheduleKind kind, string name, string initials, string expected) =>
        Assert.Equal(expected, Badge.TextFor(new ScheduleRef(kind, "1", name, initials)));

    [Fact]
    public void Color_class_per_kind() =>
        Assert.Equal(["staff", "class", "room"], new[] { ScheduleKind.Employee, ScheduleKind.Group, ScheduleKind.Resource }.Select(Badge.ClassFor));
}

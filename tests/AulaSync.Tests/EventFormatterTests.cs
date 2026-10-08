using AulaSync.Core;
using static AulaSync.Tests.TestData;

namespace AulaSync.Tests;

public class EventFormatterTests
{
    static readonly (string, string)[] AE = [("Anna Eksempel", "AE")];
    static readonly (string, string)[] AECT = [("Anna Eksempel", "AE"), ("Carl Testesen", "CT")];

    // Medarbejder: FAG | Lokale | Klasse
    [Fact] public void Employee_all() => Assert.Equal("DAN | 53 | 7A", EventFormatter.Summary(ScheduleKind.Employee, Ev("DAN", "53", ["7A"], AE)));
    [Fact] public void Employee_no_location() => Assert.Equal("Morgentilsyn | 5A", EventFormatter.Summary(ScheduleKind.Employee, Ev("Morgentilsyn", "", ["5A"], AE)));
    [Fact] public void Employee_only_title() => Assert.Equal("MØDE", EventFormatter.Summary(ScheduleKind.Employee, Ev("MØDE", "")));
    [Fact] public void Employee_two_groups() => Assert.Equal("DAN | 53 | 7A, 7B", EventFormatter.Summary(ScheduleKind.Employee, Ev("DAN", "53", ["7A", "7B"])));

    // Klasse: FAG | Lokale | INIT
    [Fact] public void Group_all() => Assert.Equal("DAN | 53 | AE", EventFormatter.Summary(ScheduleKind.Group, Ev("DAN", "53", ["7A"], AE)));
    [Fact] public void Group_two_teachers() => Assert.Equal("DAN | 53 | AE, CT", EventFormatter.Summary(ScheduleKind.Group, Ev("DAN", "53", ["7A"], AECT)));
    [Fact] public void Group_no_teacher() => Assert.Equal("MØDE | 53", EventFormatter.Summary(ScheduleKind.Group, Ev("MØDE", "53", ["7A"])));

    // Lokale: FAG | Klasse | INIT
    [Fact] public void Resource_all() => Assert.Equal("IDR | 7A | AE", EventFormatter.Summary(ScheduleKind.Resource, Ev("IDR", "GYM", ["7A"], AE)));
    [Fact] public void Resource_no_class() => Assert.Equal("MØDE | AE", EventFormatter.Summary(ScheduleKind.Resource, Ev("MØDE", "GYM", [], AE)));
    [Fact] public void Resource_only_title() => Assert.Equal("MØDE", EventFormatter.Summary(ScheduleKind.Resource, Ev("MØDE", "GYM")));

    [Fact]
    public void Substitute_gets_prefix() =>
        Assert.Equal("Vikar | MAT | 54 | 7B", EventFormatter.Summary(ScheduleKind.Employee, Ev("MAT", "54", ["7B"], [("Vera Vikar", "VV")], substitute: true)));

    [Fact]
    public void Empty_title_becomes_placeholder() =>
        Assert.Equal("Uden titel | 53", EventFormatter.Summary(ScheduleKind.Employee, Ev("", "53")));

    [Fact]
    public void Description_lists_everything()
    {
        var e = Ev("MAT", "54", ["7B"], [("Vera Vikar", "VV")], substitute: true, substituteFor: ("Anna Eksempel", "AE"));
        Assert.Equal("Vikar for: Anna Eksempel (AE)\nVera Vikar (VV)\nKlasse: 7B\nLokale: 54", EventFormatter.Description(e));
    }

    [Fact]
    public void Description_empty_for_bare_event() => Assert.Equal("", EventFormatter.Description(Ev("MØDE", "")));
}

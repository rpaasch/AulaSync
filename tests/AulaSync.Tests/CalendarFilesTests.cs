using AulaSync.Core;

namespace AulaSync.Tests;

// Kalenderfilerne hedder fx 123456-AE-medarbejder-1001.ics, så de er lette at kende fra hinanden i Stifinder og Finder.
// Adressen i kalenderprogrammet er stadig nøglen (medarbejder-1001.ics).
public class CalendarFilesTests : IDisposable
{
    static readonly ScheduleRef Anna = new(ScheduleKind.Employee, "1001", "Anna Eksempel", "AE");
    static readonly ScheduleRef SevenA = new(ScheduleKind.Group, "88231", "7A");

    readonly TempDir _dir = new();

    public void Dispose() => _dir.Dispose();

    [Theory]
    [InlineData(ScheduleKind.Employee, "1001", "Anna Eksempel", "AE", "123456-AE-medarbejder-1001.ics")]
    [InlineData(ScheduleKind.Employee, "1002", "Bo Eksempel", "", "123456-Bo-Eksempel-medarbejder-1002.ics")]
    [InlineData(ScheduleKind.Group, "88231", "7A", "", "123456-7A-klasse-88231.ics")]
    [InlineData(ScheduleKind.Group, "88232", "0.B", "", "123456-0-B-klasse-88232.ics")]
    [InlineData(ScheduleKind.Resource, "4711", "Lokale 53", "", "123456-Lokale-53-lokale-4711.ics")]
    [InlineData(ScheduleKind.Resource, "4712", " Musik/Drama (gl. sal) ", "", "123456-Musik-Drama-gl-sal-lokale-4712.ics")]
    [InlineData(ScheduleKind.Resource, "4713", "Lærerværelset", "", "123456-Lærerværelset-lokale-4713.ics")]
    [InlineData(ScheduleKind.Resource, "4714", "Århus", "", "123456-Århus-lokale-4714.ics")] // samlet til ét tegn
    [InlineData(ScheduleKind.Resource, "4715", "?!", "", "123456-lokale-4715.ics")]
    public void Name_has_institution_label_and_key(ScheduleKind kind, string id, string name, string initials, string expected) =>
        Assert.Equal(expected, CalendarFiles.Name("123456", new ScheduleRef(kind, id, name, initials)));

    [Fact]
    public void Without_institution_the_name_starts_with_the_label() =>
        Assert.Equal("AE-medarbejder-1001.ics", CalendarFiles.Name("", Anna));

    [Fact]
    public void Long_label_is_cut()
    {
        var name = CalendarFiles.Name("1", new ScheduleRef(ScheduleKind.Resource, "9", "Et meget langt lokalenavn med mange ord i sig"));
        Assert.Equal("1-Et-meget-langt-lokalenavn-med-mange-ord-lokale-9.ics", name); // højst 40 tegn, uden bindestreg sidst
    }

    // Et filsystem kan give navnet tilbage med æ, ø og å skilt ad (NFD) eller med andre store og små bogstaver; det er
    // stadig samme fil.
    [Fact]
    public void Same_name_ignores_case_and_normalisation()
    {
        Assert.True(CalendarFiles.SameName("/k/123456-Ha\u030Andværk-lokale-1.ics", "/k/123456-håndværk-lokale-1.ics"));
        Assert.False(CalendarFiles.SameName("/k/123456-7A-klasse-1.ics", "/k/klasse-1.ics"));
        Assert.False(CalendarFiles.SameName("/k/klasse-1.ics", null));
    }

    // Filen findes ud fra nøglen sidst i navnet; filer fra 3.0.0 hedder bare nøglen. Nyeste først.
    [Fact]
    public void Find_matches_the_key_at_the_end_newest_first()
    {
        var old = _dir.File("klasse-88231.ics");
        var renamed = _dir.File("123456-7A-klasse-88231.ics");
        File.WriteAllText(old, "");
        File.WriteAllText(renamed, "");
        File.SetLastWriteTimeUtc(old, new DateTime(2026, 10, 1, 0, 0, 0, DateTimeKind.Utc));
        File.WriteAllText(_dir.File("123456-7A-klasse-188231.ics"), "");   // et andet id, der ender på samme cifre
        File.WriteAllText(_dir.File("123456-7A-lokale-88231.ics"), "");    // en anden slags
        File.WriteAllText(_dir.File("klasse-88231.ics.0f3a.tmp"), "");     // AtomicFile undervejs
        File.WriteAllText(_dir.File("abonnementer.json"), "");

        Assert.Equal([renamed, old], CalendarFiles.Find(_dir.Path, SevenA.Key));
        Assert.Empty(CalendarFiles.Find(_dir.Path, Anna.Key));
        Assert.Empty(CalendarFiles.Find(_dir.File("findes-ikke"), SevenA.Key));
    }
}

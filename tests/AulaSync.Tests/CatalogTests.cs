using AulaSync.Core;

namespace AulaSync.Tests;

public class NaturalComparerTests
{
    [Fact]
    public void Classes_sort_by_number_then_letter() =>
        Assert.Equal(["0A", "1A", "1B", "2A", "9C", "10A", "10B"],
            new[] { "10B", "1B", "0A", "9C", "10A", "2A", "1A" }.Order(NaturalComparer.Danish));

    [Fact]
    public void Rooms_with_text_and_numbers() =>
        Assert.Equal(["Lokale 2", "Lokale 10", "Lokale 53", "Æblerummet", "Ølstuen"],
            new[] { "Lokale 53", "Ølstuen", "Lokale 10", "Æblerummet", "Lokale 2" }.Order(NaturalComparer.Danish));

    [Fact]
    public void Case_is_ignored() => Assert.Equal(0, NaturalComparer.Danish.Compare("gym", "GYM"));

    [Fact]
    public void Leading_zeros_compare_equal_numerically() => Assert.Equal(0, NaturalComparer.Danish.Compare("Lokale 07", "Lokale 7"));
}

public class ScheduleOrderingTests
{
    [Fact]
    public void Employees_by_initials_then_groups_then_resources()
    {
        ScheduleRef[] input =
        [
            new(ScheduleKind.Resource, "1", "Lokale 10"),
            new(ScheduleKind.Group, "2", "10A"),
            new(ScheduleKind.Employee, "3", "Bo Testesen", "BT"),
            new(ScheduleKind.Resource, "4", "Lokale 2"),
            new(ScheduleKind.Employee, "5", "Anna Eksempel", "AE"),
            new(ScheduleKind.Group, "6", "1A"),
        ];
        Assert.Equal(["5", "3", "6", "2", "4", "1"], ScheduleOrdering.Sort(input).Select(s => s.Id));
    }
}

public class ScheduleSearchTests
{
    static readonly ScheduleRef Anna = new(ScheduleKind.Employee, "1", "Anna Eksempel", "AE");
    static readonly ScheduleRef Bo = new(ScheduleKind.Employee, "2", "Bo Testesen", "BT");
    static readonly ScheduleRef SevenA = new(ScheduleKind.Group, "3", "7A");
    static readonly ScheduleRef Gym = new(ScheduleKind.Resource, "4", "Gymnastiksal");
    static readonly IReadOnlyList<ScheduleRef> All = [Anna, Bo, SevenA, Gym];
    static readonly IReadOnlySet<string> None = new HashSet<string>();

    [Theory]
    [InlineData("anna", true)]
    [InlineData("AE", true)]
    [InlineData("ae anna", true)]
    [InlineData("eksempel", true)]
    [InlineData("xyz", false)]
    public void Matches_name_and_initials(string query, bool expected) => Assert.Equal(expected, ScheduleSearch.Matches(Anna, query));

    [Fact]
    public void Empty_query_shows_suggestion_and_everything_in_groups()
    {
        var rows = ScheduleSearch.BuildRows(All, "", CatalogFilter.All, Anna, None);
        Assert.Equal(
        [
            new CatalogHeader("Forslag"), new CatalogEntry(Anna, false, true),
            new CatalogHeader("Medarbejdere"), new CatalogEntry(Anna, false, false), new CatalogEntry(Bo, false, false),
            new CatalogHeader("Klasser"), new CatalogEntry(SevenA, false, false),
            new CatalogHeader("Lokaler"), new CatalogEntry(Gym, false, false),
        ], rows);
    }

    [Fact]
    public void Suggestion_is_hidden_when_already_subscribed()
    {
        var rows = ScheduleSearch.BuildRows(All, "", CatalogFilter.All, Anna, new HashSet<string> { Anna.Key });
        Assert.DoesNotContain(new CatalogHeader("Forslag"), rows);
        Assert.Contains(new CatalogEntry(Anna, true, false), rows);
    }

    [Fact]
    public void Query_and_filter_narrow_the_list()
    {
        Assert.Equal([new CatalogHeader("Klasser"), new CatalogEntry(SevenA, false, false)],
            ScheduleSearch.BuildRows(All, "7", CatalogFilter.All, Anna, None));
        Assert.Equal([new CatalogHeader("Lokaler"), new CatalogEntry(Gym, false, false)],
            ScheduleSearch.BuildRows(All, "", CatalogFilter.Resources, Anna, None));
        Assert.Empty(ScheduleSearch.BuildRows(All, "findes ikke", CatalogFilter.All, Anna, None));
    }
}

public class ScheduleCatalogTests
{
    sealed class FakeDirectory : IAulaDirectory
    {
        public int Calls;
        public bool Fail;
        public Task<IReadOnlyList<Employee>> GetEmployeesAsync(CancellationToken ct)
        {
            Calls++;
            if (Fail) throw new AulaException("nede");
            return Task.FromResult<IReadOnlyList<Employee>>([new("1", "Bo Testesen", "BT", "teacher"), new("2", "Anna Eksempel", "AE", "teacher")]);
        }
        public Task<IReadOnlyList<NamedItem>> GetGroupsAsync(CancellationToken ct) => Task.FromResult<IReadOnlyList<NamedItem>>([new("10", "10A"), new("11", "1A")]);
        public Task<IReadOnlyList<NamedItem>> GetResourcesAsync(CancellationToken ct) => Task.FromResult<IReadOnlyList<NamedItem>>([new("20", "Lokale 53")]);
    }

    [Fact]
    public async Task Loads_once_maps_and_sorts()
    {
        var dir = new FakeDirectory();
        var catalog = new ScheduleCatalog(dir);
        var all = await catalog.GetAllAsync();
        await catalog.GetAllAsync();
        Assert.Equal(1, dir.Calls);
        Assert.Equal(["AE Anna Eksempel", "BT Bo Testesen", "1A", "10A", "Lokale 53"], all.Select(s => s.CalendarName));
        Assert.Equal(["teacher", "teacher", "", "", ""], all.Select(s => s.Role));
    }

    // Når listen er hentet, får kalderen den (AulaSync opdaterer så rollerne på valgte skemaer).
    [Fact]
    public async Task Reports_the_loaded_list_but_not_a_failure()
    {
        var dir = new FakeDirectory { Fail = true };
        var reported = new List<IReadOnlyList<ScheduleRef>>();
        var catalog = new ScheduleCatalog(dir, reported.Add);
        await Assert.ThrowsAsync<AulaException>(() => catalog.GetAllAsync());
        Assert.Empty(reported);
        dir.Fail = false;
        var all = await catalog.GetAllAsync();
        Assert.Same(all, Assert.Single(reported));
    }

    [Fact]
    public async Task Failure_is_not_cached()
    {
        var dir = new FakeDirectory { Fail = true };
        var catalog = new ScheduleCatalog(dir);
        await Assert.ThrowsAsync<AulaException>(() => catalog.GetAllAsync());
        dir.Fail = false;
        Assert.Equal(5, (await catalog.GetAllAsync()).Count);
    }
}

public class CalendarAppsTests
{
    [Theory]
    [InlineData(CalendarApp.AppleCalendar, "Tilføj til Kalender", MainActionKind.Subscribe)]
    [InlineData(CalendarApp.OutlookClassic, "Tilføj til Outlook", MainActionKind.Subscribe)]
    [InlineData(CalendarApp.OutlookImport, "Importér…", MainActionKind.Import)]
    [InlineData(CalendarApp.Other, "Kopiér adresse", MainActionKind.CopyAddress)]
    public void Main_action_per_app(CalendarApp app, string label, MainActionKind kind) =>
        Assert.Equal(new MainAction(label, kind), CalendarApps.ActionFor(app));

    [Fact]
    public void Default_is_apple_calendar_on_mac_and_nothing_on_windows()
    {
        Assert.Equal(CalendarApp.AppleCalendar, CalendarApps.DefaultFor(isMac: true));
        Assert.Null(CalendarApps.DefaultFor(isMac: false));
    }

    // Teksterne på valgkortene i første start og i Indstillinger (UI-mockup).
    [Theory]
    [InlineData(CalendarApp.AppleCalendar, "Apple Kalender", true, "Kalender-appen på denne Mac. Skemaerne opdateres af sig selv, så længe AulaSync kører.")]
    [InlineData(CalendarApp.OutlookClassic, "Outlook (klassisk)", true, "Den klassiske Outlook til Windows. Skemaerne tilføjes som internetkalendere og opdateres af sig selv, så længe AulaSync kører.")]
    [InlineData(CalendarApp.OutlookImport, "Ny Outlook, Outlook til Mac eller web", false, "Henter kalendere gennem Microsofts servere, som ikke kan nå denne computer. Du importerer en fil, og AulaSync siger til, når skemaet er ændret, så du kan importere igen.")]
    [InlineData(CalendarApp.Other, "Andet program", true, "Fx Thunderbird. Kopiér en .ics-adresse, og indsæt den i et kalenderprogram på denne computer. Opdateres, så længe AulaSync kører.")]
    public void Card_texts(CalendarApp app, string title, bool automatic, string description)
    {
        Assert.Equal(title, CalendarApps.Title(app));
        Assert.Equal(automatic, CalendarApps.UpdatesAutomatically(app));
        Assert.Equal(description, CalendarApps.Description(app));
    }

    [Fact]
    public void Mode_label()
    {
        Assert.Equal("Opdateres automatisk", CalendarApps.ModeLabel(CalendarApp.AppleCalendar));
        Assert.Equal("Øjebliksbillede", CalendarApps.ModeLabel(CalendarApp.OutlookImport));
    }
}

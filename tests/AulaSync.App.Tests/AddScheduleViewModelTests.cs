using Avalonia.Controls;
using Avalonia.Headless.XUnit;
using Avalonia.VisualTree;
using AulaSync.Core;
using static AulaSync.Tests.TestData;

namespace AulaSync.App.Tests;

public class AddScheduleViewModelTests : IDisposable
{
    static readonly ScheduleRef Anna = new(ScheduleKind.Employee, "1001", "Anna Eksempel", "AE");

    readonly TestHost _host = new();
    readonly FakeDirectory _directory = new();
    ScheduleCatalog? _catalog;

    public AddScheduleViewModelTests()
    {
        _catalog = new ScheduleCatalog(_directory);
        _host.Sync.SetClient(_host.Client);
    }

    public void Dispose() => _host.Dispose();

    AddScheduleViewModel Create() => new(() => _catalog, () => Anna, _host.Sync, _host.Log, a => a());

    static List<string> Names(AddScheduleViewModel vm) =>
        vm.Rows.Select(r => r switch { HeaderRow h => $"# {h.Title}", CatalogEntryViewModel e => e.Name, _ => "?" }).ToList();

    [Fact]
    public async Task Empty_search_shows_suggestion_then_everything_in_groups()
    {
        var vm = Create();
        await vm.LoadAsync();
        Assert.Equal(["# Forslag", "AE Anna Eksempel", "# Medarbejdere", "AE Anna Eksempel", "BT Bo Testesen",
            "# Klasser", "1A", "7A", "10A", "# Lokaler", "Gymnastiksal", "Lokale 53"], Names(vm));
        var suggestion = (CatalogEntryViewModel)vm.Rows[1];
        Assert.Equal("Dit eget skema", suggestion.Sub);
        Assert.Equal("Lærer", ((CatalogEntryViewModel)vm.Rows[3]).Sub); // rollen fra Aula
        Assert.Equal("0 skemaer valgt", vm.CountText);
    }

    // Rækkerne under hver overskrift tegnes som ét kort med runde hjørner øverst og nederst (UI-mockuppen).
    [Fact]
    public async Task Rows_know_their_place_in_each_group()
    {
        var vm = Create();
        await vm.LoadAsync();
        var marks = vm.Rows.Select(r => r is CatalogEntryViewModel e ? $"{e.Name}{(e.IsFirst ? " første" : "")}{(e.IsLast ? " sidste" : "")}" : "#").ToList();
        Assert.Equal(["#", "AE Anna Eksempel første sidste", "#", "AE Anna Eksempel første", "BT Bo Testesen sidste",
            "#", "1A første", "7A", "10A sidste", "#", "Gymnastiksal første", "Lokale 53 sidste"], marks);
    }

    [Fact]
    public async Task Search_and_filter_narrow_the_list()
    {
        var vm = Create();
        await vm.LoadAsync();
        vm.Query = "7";
        Assert.Equal(["# Klasser", "7A"], Names(vm));
        vm.Query = "";
        vm.IsResources = true;
        Assert.Equal(CatalogFilter.Resources, vm.Filter);
        Assert.False(vm.IsAll);
        Assert.Equal(["# Lokaler", "Gymnastiksal", "Lokale 53"], Names(vm));
        vm.Query = "findes ikke";
        Assert.Equal("Ingen resultater for \"findes ikke\".", vm.EmptyText);
    }

    [Fact]
    public async Task Add_marks_row_counts_and_subscribes()
    {
        var vm = Create();
        await vm.LoadAsync();
        var sevenA = vm.Rows.OfType<CatalogEntryViewModel>().Single(e => e.Name == "7A");
        sevenA.AddCommand.Execute(null);
        Assert.True(sevenA.IsAdded);
        Assert.Equal("1 skema valgt", vm.CountText);
        await WaitUntil(() => _host.Store.Load().Any(s => s.Key == "klasse-7"));

        var own = vm.Rows.OfType<CatalogEntryViewModel>().First();
        own.AddCommand.Execute(null);
        Assert.Equal("2 skemaer valgt", vm.CountText);
        Assert.All(vm.Rows.OfType<CatalogEntryViewModel>().Where(e => e.Schedule.Key == Anna.Key), e => Assert.True(e.IsAdded));
    }

    [Fact]
    public async Task Already_subscribed_shows_added_and_hides_suggestion()
    {
        await _host.Sync.AddAsync(Anna, default);
        var vm = Create();
        await vm.LoadAsync();
        Assert.DoesNotContain("# Forslag", Names(vm));
        Assert.True(vm.Rows.OfType<CatalogEntryViewModel>().Single(e => e.Name == "AE Anna Eksempel").IsAdded);
        Assert.Equal("1 skema valgt", vm.CountText);
    }

    [Fact]
    public async Task Failure_shows_error_and_retry_works()
    {
        _directory.Fail = true;
        var vm = Create();
        await vm.LoadAsync();
        Assert.Equal("Listen kunne ikke hentes: nede", vm.Error);
        Assert.Equal(["# Forslag", "AE Anna Eksempel"], Names(vm)); // eget skema kan stadig tilføjes

        _directory.Fail = false;
        await vm.LoadCommand.ExecuteAsync(null);
        Assert.Null(vm.Error);
        Assert.NotEmpty(vm.Rows);
    }

    [Fact]
    public async Task Not_logged_in()
    {
        _catalog = null;
        var vm = Create();
        await vm.LoadAsync();
        Assert.Equal("Du er ikke logget ind i Aula.", vm.Error);
    }

    [AvaloniaFact]
    public async Task Window_shows_rows_and_count()
    {
        var vm = Create();
        var window = new AddScheduleWindow(vm);
        window.Show();
        await vm.LoadAsync();
        Avalonia.Threading.Dispatcher.UIThread.RunJobs();
        var texts = window.GetVisualDescendants().OfType<TextBlock>().Select(t => t.Text).ToList();
        Assert.Contains("Dit eget skema", texts);
        Assert.Contains("0 skemaer valgt", texts);
        Assert.Contains(window.GetVisualDescendants().OfType<Button>(), b => Avalonia.Automation.AutomationProperties.GetName(b) == "Tilføj 7A");
    }
}

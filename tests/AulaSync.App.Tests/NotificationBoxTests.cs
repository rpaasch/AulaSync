using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Headless.XUnit;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Avalonia.VisualTree;
using Microsoft.Extensions.Time.Testing;

namespace AulaSync.App.Tests;

public class NotificationBoxTests : IDisposable
{
    readonly TestHost _host = new();
    readonly FakeTimeProvider _time = new();

    public void Dispose() => _host.Dispose();

    NotificationBox Create() => new(new FakePlatform(), _host.Log, _time);

    static Button Find(Window window, string name) =>
        window.GetVisualDescendants().OfType<Button>().Single(b => AutomationProperties.GetName(b) == name);

    static void Click(Button button) => button.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

    void Advance(int seconds)
    {
        _time.Advance(TimeSpan.FromSeconds(seconds));
        Dispatcher.UIThread.RunJobs();
    }

    [AvaloniaFact]
    public void Shows_text_on_top_without_taking_focus()
    {
        var box = Create();
        box.Show(Relogin.ExpiredText);
        var window = Assert.IsType<Window>(box.Current);
        Assert.True(window.IsVisible);
        Assert.True(window.Topmost);
        Assert.False(window.ShowActivated);
        Assert.False(window.ShowInTaskbar);
        Assert.Contains(window.GetVisualDescendants().OfType<TextBlock>(), t => t.Text == Relogin.ExpiredText);
        Assert.Contains($"Notifikation: {Relogin.ExpiredText}", File.ReadAllText(_host.Paths.Log));
    }

    [AvaloniaFact]
    public void Click_on_message_closes_and_reports_click()
    {
        var box = Create();
        int clicks = 0, dismissed = 0;
        box.Clicked += () => clicks++;
        box.Dismissed += () => dismissed++;
        box.Show(Relogin.ExpiredText);
        var window = box.Current!;
        Click(Find(window, Relogin.ExpiredText));
        Assert.Equal(1, clicks);
        Assert.Equal(0, dismissed);
        Assert.Null(box.Current);
        Assert.False(window.IsVisible);
    }

    // En besked kan have sin egen handling (fx åbn hovedvinduet); så går klikket ikke til login (Clicked).
    [AvaloniaFact]
    public void Click_runs_the_messages_own_action()
    {
        var box = Create();
        int clicks = 0, opened = 0;
        box.Clicked += () => clicks++;
        box.Show("7A er ændret siden import. Klik her for at importere igen.", () => opened++);
        Click(Find(box.Current!, "7A er ændret siden import. Klik her for at importere igen."));
        Assert.Equal(1, opened);
        Assert.Equal(0, clicks);
        Assert.Null(box.Current);
    }

    // ✕ giver Dismissed (på Mac skjuler AppHost så appen, så det forrige program får fokus igen).
    [AvaloniaFact]
    public void Close_button_closes_and_reports_dismissal()
    {
        var box = Create();
        int clicks = 0, dismissed = 0;
        box.Clicked += () => clicks++;
        box.Dismissed += () => dismissed++;
        box.Show(Relogin.ExpiredText);
        Click(Find(box.Current!, "Luk"));
        Assert.Equal(0, clicks);
        Assert.Equal(1, dismissed);
        Assert.Null(box.Current);
    }

    [AvaloniaFact]
    public void Closes_itself_after_20_seconds_and_a_new_message_replaces_the_old()
    {
        var box = Create();
        var dismissed = 0;
        box.Dismissed += () => dismissed++;
        box.Show("Første");
        var first = box.Current!;
        Advance(10);
        box.Show("Anden");
        Assert.False(first.IsVisible);
        Advance(10);
        Assert.Contains(box.Current!.GetVisualDescendants().OfType<TextBlock>(), t => t.Text == "Anden");
        Advance(9);
        Assert.NotNull(box.Current);
        Advance(1);
        Assert.Null(box.Current);
        Assert.Equal(0, dismissed);
    }

    // Øverst til højre under menulinjen på Mac, nederst til højre over proceslinjen på Windows; 12 punkter fra kanten.
    [Theory]
    [InlineData(true, 1, 1568, 37)]
    [InlineData(false, 1, 1568, 948)]
    [InlineData(false, 2, 3136, 1896)]
    public void Corner_follows_the_platform(bool isMac, double scaling, int x, int y)
    {
        var area = isMac ? new PixelRect(0, 25, (int)(1920 * scaling), (int)(1055 * scaling)) : new PixelRect(0, 0, (int)(1920 * scaling), (int)(1040 * scaling));
        Assert.Equal(new PixelPoint(x, y), NotificationBox.Corner(area, new Size(340, 80), scaling, top: isMac));
    }
}

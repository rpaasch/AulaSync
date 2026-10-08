using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Platform;
using Avalonia.Threading;
using AulaSync.Core;

namespace AulaSync.App;

// AulaSyncs egen notifikationsboks: et lille vindue i skærmens hjørne (øverst til højre på Mac, nederst til højre på
// Windows) over andre vinduer, uden at tage fokus. Klik på beskeden giver Clicked, ✕ lukker og giver Dismissed, og ellers
// lukker boksen sig selv efter 20 sekunder. Den kræver ingen tilladelse fra styresystemet. Mac: boksen må vises i alle
// Spaces og over programmer i fuld skærm.
public sealed class NotificationBox(IPlatform platform, FileLog log, TimeProvider time) : INotifier
{
    public static readonly TimeSpan Duration = TimeSpan.FromSeconds(20);
    const double Width = 340, Margin = 12;

    Window? _window;
    ITimer? _timer;

    public event Action? Clicked;

    public event Action? Dismissed;

    // Den viste boks, eller null.
    public Window? Current => _window;

    public void Show(string text, Action? onClick = null)
    {
        log.Info($"Notifikation: {text}");
        Dismiss();

        var message = new Button
        {
            Classes = { "ghost" },
            HorizontalAlignment = HorizontalAlignment.Stretch,
            HorizontalContentAlignment = HorizontalAlignment.Left,
            Padding = new Thickness(14, 12, 4, 12),
            Content = new StackPanel
            {
                Spacing = 2,
                Children =
                {
                    new TextBlock { Text = "AulaSync", FontWeight = FontWeight.SemiBold },
                    new TextBlock { Text = text, TextWrapping = TextWrapping.Wrap },
                },
            },
        };
        AutomationProperties.SetName(message, text);
        var close = new Button { Classes = { "ghost" }, Content = "✕", VerticalAlignment = VerticalAlignment.Top, Margin = new Thickness(0, 6, 6, 0) };
        AutomationProperties.SetName(close, "Luk");
        var grid = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto"), Children = { message, close } };
        Grid.SetColumn(close, 1);
        var border = new Border { BorderThickness = new Thickness(1), Child = grid };
        border.Bind(Border.BorderBrushProperty, border.GetResourceObservable("LineBrush"));

        var window = new Window
        {
            Title = "AulaSync",
            Width = Width,
            SizeToContent = SizeToContent.Height,
            CanResize = false,
            Topmost = true,
            ShowActivated = false,
            ShowInTaskbar = false,
            WindowDecorations = WindowDecorations.None,
            WindowStartupLocation = WindowStartupLocation.Manual,
            Opacity = 0, // vises først, når den står i hjørnet
            Content = border,
        };
        window.Opened += (_, _) =>
        {
            if (window.Screens.Primary is { } screen)
                window.Position = Corner(screen.WorkingArea, window.ClientSize, screen.Scaling, top: platform.IsMac);
            window.Opacity = 1;
        };
        message.Click += (_, _) =>
        {
            Dismiss();
            if (onClick is not null) onClick();
            else Clicked?.Invoke();
        };
        close.Click += (_, _) => { Dismiss(); Dismissed?.Invoke(); };

        if (window.TryGetPlatformHandle() is IMacOSTopLevelPlatformHandle mac) MacDock.ShowOnAllSpaces(mac.NSWindow);
        MacDock.UnhideWithoutActivation(); // efter ✕ er appen skjult; ellers ville boksen ikke kunne ses

        _window = window;
        window.Show();
        _timer = time.CreateTimer(_ => Dispatcher.UIThread.Post(() => { if (_window == window) Dismiss(); }),
            null, Duration, Timeout.InfiniteTimeSpan);
    }

    public void Dismiss()
    {
        _timer?.Dispose();
        _timer = null;
        var window = _window;
        _window = null;
        window?.Close();
    }

    // Hjørnet i skærmens arbejdsområde (uden menulinje og proceslinje), Margin punkter fra kanten.
    public static PixelPoint Corner(PixelRect area, Size size, double scaling, bool top)
    {
        var pixels = PixelSize.FromSize(size, scaling);
        var margin = (int)Math.Round(Margin * scaling);
        return new PixelPoint(area.Right - pixels.Width - margin, top ? area.Y + margin : area.Bottom - pixels.Height - margin);
    }
}

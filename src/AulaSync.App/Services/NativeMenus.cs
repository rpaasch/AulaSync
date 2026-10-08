using Avalonia.Controls;
using Avalonia.Input;

namespace AulaSync.App;

// Ikonets menu og Mac-appens menu. Avalonia på Mac låser sig til den første NativeMenu, et ikon eller appen får:
// en ny menu til ikonet senere får appen til at gå ned ("The menu being updated does not match"), og en app-menu sat
// efter opstarten ignoreres. Derfor har ikonet én menu, hvis punkter udskiftes, og app-menuen sættes i App.Initialize.
public static class NativeMenus
{
    // Udskift punkterne i ikonets menu. Genvejstaster (⌘) kun på Mac.
    public static void Fill(NativeMenu menu, IEnumerable<TrayItem> items, Action<TrayCommand> run, bool mac)
    {
        menu.Items.Clear();
        foreach (var item in items)
        {
            if (item == TrayItem.Separator) { menu.Items.Add(new NativeMenuItemSeparator()); continue; }
            var entry = new NativeMenuItem(item.Label) { IsEnabled = item.Command is not null };
            if (item.Shortcut is { } key && mac) entry.Gesture = Gesture(key);
            if (item.Command is { } command) entry.Click += (_, _) => run(command);
            menu.Items.Add(entry);
        }
    }

    // Mac-appens menu på dansk. Avalonias egne punkter (engelske) er slået fra i Program.BuildAvaloniaApp.
    public static NativeMenu AppMenu(Action openSettings, Action quit)
    {
        var menu = new NativeMenu();
        void Add(string header, KeyGesture? gesture, Action click)
        {
            var item = new NativeMenuItem(header) { Gesture = gesture };
            item.Click += (_, _) => click();
            menu.Items.Add(item);
        }
        Add("Indstillinger…", Gesture(','), openSettings);
        menu.Items.Add(new NativeMenuItemSeparator());
        Add("Skjul AulaSync", Gesture('h'), MacDock.HideApp);
        Add("Skjul andre", new KeyGesture(Key.H, KeyModifiers.Meta | KeyModifiers.Alt), MacDock.HideOthers);
        Add("Vis alle", null, MacDock.ShowAll);
        menu.Items.Add(new NativeMenuItemSeparator());
        Add("Afslut AulaSync", Gesture('q'), quit);
        return menu;
    }

    public static KeyGesture Gesture(char key) =>
        new(key == ',' ? Key.OemComma : Enum.Parse<Key>(char.ToUpperInvariant(key).ToString()), KeyModifiers.Meta);
}

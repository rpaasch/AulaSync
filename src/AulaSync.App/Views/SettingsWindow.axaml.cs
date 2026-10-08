using Avalonia.Controls;
using Avalonia.Input;

namespace AulaSync.App;

// Mac: eget vindue (⌘,). Windows: dialog (Ctrl+,).
public partial class SettingsWindow : Window
{
    public SettingsWindow() => InitializeComponent();

    public SettingsWindow(SettingsViewModel vm) : this()
    {
        DataContext = vm;
        KeyDown += (_, e) =>
        {
            if (e.Key == Key.Escape || (e.Key == Key.W && e.KeyModifiers == KeyModifiers.Meta)) Close();
        };
    }
}

using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;

namespace AulaSync.App;

// Mac: spec ønsker et sheet; Avalonia har ingen sheets, så det er en modal dialog centreret over hovedvinduet på begge platforme.
public partial class AddScheduleWindow : Window
{
    public AddScheduleWindow() => InitializeComponent();

    public AddScheduleWindow(AddScheduleViewModel vm) : this()
    {
        DataContext = vm;
        Opened += async (_, _) => await vm.LoadAsync();
        KeyDown += (_, e) => { if (e.Key == Key.Escape) Close(); };
    }

    void Done(object? sender, RoutedEventArgs e) => Close();
}

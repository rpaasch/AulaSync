using Avalonia.Controls;
using Avalonia.Interactivity;

namespace AulaSync.App;

public partial class OutlookFallbackWindow : Window
{
    public OutlookFallbackWindow() => InitializeComponent();

    public OutlookFallbackWindow(OutlookFallbackViewModel vm) : this() => DataContext = vm;

    void Ok(object? sender, RoutedEventArgs e) => Close();
}

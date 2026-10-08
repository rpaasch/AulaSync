using Avalonia.Controls;
using Avalonia.Interactivity;

namespace AulaSync.App;

public partial class ConfirmLogoutWindow : Window
{
    public ConfirmLogoutWindow()
    {
        InitializeComponent();
        ButtonOrder.Apply(Buttons, Primary);
    }

    void Confirm(object? sender, RoutedEventArgs e) => Close(true);

    void Cancel(object? sender, RoutedEventArgs e) => Close(false);
}

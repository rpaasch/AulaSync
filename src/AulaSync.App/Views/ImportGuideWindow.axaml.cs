using Avalonia.Controls;
using Avalonia.Interactivity;

namespace AulaSync.App;

// Lukker med true, når brugeren klikker Færdig; så markeres skemaet som importeret (dato + hash).
public partial class ImportGuideWindow : Window
{
    public ImportGuideWindow() => InitializeComponent();

    public ImportGuideWindow(ImportGuideViewModel vm) : this()
    {
        DataContext = vm;
        ButtonOrder.Apply(Buttons, Primary);
    }

    void Done(object? sender, RoutedEventArgs e) => Close(true);

    void Cancel(object? sender, RoutedEventArgs e) => Close(false);
}

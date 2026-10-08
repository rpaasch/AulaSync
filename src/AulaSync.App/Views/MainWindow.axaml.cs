using Avalonia.Controls;
using Avalonia.Input;

namespace AulaSync.App;

public partial class MainWindow : Window
{
    public MainWindow() => InitializeComponent();

    public MainWindow(MainViewModel vm) : this()
    {
        DataContext = vm;
        // Platformens genveje: Mac ⌘N, ⌘R, ⌘W · Windows Ctrl+N, F5 (spec §3.2).
        var mac = OperatingSystem.IsMacOS();
        KeyBindings.Add(new KeyBinding { Gesture = new KeyGesture(Key.N, mac ? KeyModifiers.Meta : KeyModifiers.Control), Command = vm.AddScheduleCommand });
        KeyBindings.Add(new KeyBinding { Gesture = mac ? new KeyGesture(Key.R, KeyModifiers.Meta) : new KeyGesture(Key.F5), Command = vm.RefreshCommand });
        if (mac) KeyBindings.Add(new KeyBinding { Gesture = new KeyGesture(Key.W, KeyModifiers.Meta), Command = new CloseCommand(this) });
    }

    sealed class CloseCommand(Window window) : System.Windows.Input.ICommand
    {
        public event EventHandler? CanExecuteChanged { add { } remove { } }
        public bool CanExecute(object? parameter) => true;
        public void Execute(object? parameter) => window.Close();
    }
}

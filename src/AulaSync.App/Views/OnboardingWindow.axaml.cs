using Avalonia.Controls;

namespace AulaSync.App;

public partial class OnboardingWindow : Window
{
    public OnboardingWindow() => InitializeComponent();

    // loginView: LoginView med den persistente profil (Task 13). Uden skærm/webview i testene: en tom kontrol.
    // Den sættes først ind ved trin 2: på Windows starter WebView2, så snart den er i vinduet, også skjult, og lukkede man
    // velkomsten, mens WebView2 startede, fejlede det (E_ABORT).
    public OnboardingWindow(OnboardingViewModel vm, Control loginView) : this()
    {
        DataContext = vm;
        void AttachLogin() { if (vm.IsLogin) LoginHost.Content ??= loginView; }
        AttachLogin();
        vm.PropertyChanged += (_, _) => AttachLogin();
        ButtonOrder.Apply(Buttons, Primary);
        vm.Finished += Close;
    }
}

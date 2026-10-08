using Avalonia.Controls;

namespace AulaSync.App;

// Platformens knaprækkefølge i dialoger: Mac har hovedknappen yderst til højre, Windows har den først (spec §3.2).
public static class ButtonOrder
{
    public static void Apply(Panel buttons, Control primary)
    {
        buttons.Children.Remove(primary);
        if (OperatingSystem.IsMacOS()) buttons.Children.Add(primary);
        else buttons.Children.Insert(0, primary);
    }
}

using Avalonia;
using Avalonia.Controls;

namespace AulaSync.App;

// Et vindue må ikke være højere end skærmens arbejdsområde. Ellers sætter CenterScreen titellinjen over skærmen (fx login
// ved 1920x1080 og 150 %), og vinduet kan hverken flyttes eller lukkes med musen.
public static class WindowFit
{
    const double TitleBar = 48; // titellinje og lidt luft, i punkter

    public static double Height(double wanted, PixelRect workArea, double scaling) =>
        Math.Min(wanted, workArea.Height / scaling - TitleBar);

    public static void Apply(Window window, PixelRect workArea, double scaling)
    {
        var height = Height(window.Height, workArea, scaling);
        window.MinHeight = Math.Min(window.MinHeight, height);
        window.Height = height;
    }

    // Før vinduet vises første gang.
    public static void ToScreen(Window window)
    {
        if (window.Screens.Primary is { } screen) Apply(window, screen.WorkingArea, screen.Scaling);
    }

    // Et vindue, der tilpasser højden til indholdet (Indstillinger), bliver højst skærmens højde; indholdet kan så rulles.
    public static void Cap(Window window, PixelRect workArea, double scaling) =>
        window.MaxHeight = Height(double.PositiveInfinity, workArea, scaling);

    public static void CapToScreen(Window window)
    {
        if (window.Screens.Primary is { } screen) Cap(window, screen.WorkingArea, screen.Scaling);
    }
}

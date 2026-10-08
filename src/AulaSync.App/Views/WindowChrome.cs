using Avalonia.Controls;
using Avalonia.Media;

namespace AulaSync.App;

// Windows 11: Mica-baggrund, hvor det er muligt (afprøvet i spiken, punkt 6e). Ellers systemets almindelige baggrund.
public static class WindowChrome
{
    public static void Apply(Window window)
    {
        if (!OperatingSystem.IsWindows()) return;
        window.TransparencyLevelHint = [WindowTransparencyLevel.Mica, WindowTransparencyLevel.None];
        window.Opened += (_, _) =>
        {
            if (window.ActualTransparencyLevel == WindowTransparencyLevel.Mica) window.Background = Brushes.Transparent;
        };
    }
}

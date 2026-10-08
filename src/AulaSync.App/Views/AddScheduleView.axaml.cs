using Avalonia.Controls;

namespace AulaSync.App;

public partial class AddScheduleView : UserControl
{
    public AddScheduleView()
    {
        InitializeComponent();
        // Søgefeltet har fokus, så man kan skrive med det samme.
        AttachedToVisualTree += (_, _) => Search.Focus();
    }
}

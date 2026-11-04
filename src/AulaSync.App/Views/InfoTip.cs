using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;

namespace AulaSync.App;

// Et lille ⓘ ved en indstilling (Indstillinger): et klik, Enter eller mellemrum viser hjælpeteksten (Text) i en lille
// boks under det. Udseendet står i temaet (ControlTheme for InfoTip).
public sealed class InfoTip : Button
{
    public static readonly StyledProperty<string> TextProperty = AvaloniaProperty.Register<InfoTip, string>(nameof(Text), "");

    public string Text
    {
        get => GetValue(TextProperty);
        set => SetValue(TextProperty, value);
    }

    public InfoTip()
    {
        var text = new TextBlock { TextWrapping = TextWrapping.Wrap, MaxWidth = 300, FontSize = 12.5 };
        text.Bind(TextBlock.TextProperty, this.GetObservable(TextProperty));
        Flyout = new Flyout { Content = text, Placement = PlacementMode.BottomEdgeAlignedLeft };
    }
}

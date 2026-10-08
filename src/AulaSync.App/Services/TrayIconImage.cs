using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using Avalonia.Media.Imaging;

namespace AulaSync.App;

public enum TrayState { Normal, Updating, Attention }

// Ikonet tegnes i kode: et lille kalenderblad. Sort på gennemsigtig baggrund på Mac (skabelon-ikon),
// i accentfarve på Windows. "Opdaterer" = prik midt i bladet, "kræver handling" = prik i hjørnet.
public static class TrayIconImage
{
    public static Bitmap Render(TrayState state, bool template)
    {
        var bitmap = new RenderTargetBitmap(new PixelSize(64, 64), new Vector(96, 96));
        using (var ctx = bitmap.CreateDrawingContext())
        {
            IBrush ink = template ? Brushes.Black : new SolidColorBrush(Color.Parse("#2f6fde"));
            var pen = new Pen(ink, 6);
            ctx.DrawRectangle(null, pen, new RoundedRect(new Rect(8, 12, 46, 44), 8));
            ctx.DrawLine(pen, new Point(8, 26), new Point(54, 26));
            ctx.DrawLine(pen, new Point(21, 6), new Point(21, 16));
            ctx.DrawLine(pen, new Point(41, 6), new Point(41, 16));
            if (state == TrayState.Updating) ctx.DrawEllipse(ink, null, new Point(31, 41), 6, 6);
            if (state == TrayState.Attention)
                ctx.DrawEllipse(template ? Brushes.Black : new SolidColorBrush(Color.Parse("#b26a00")), null, new Point(52, 12), 11, 11);
        }
        return bitmap;
    }

    public static WindowIcon Create(TrayState state, bool template) => new(Render(state, template));
}

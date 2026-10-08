using System.Net;
using System.Text.RegularExpressions;

namespace AulaSync.Core;

public static partial class HtmlText
{
    [GeneratedRegex(@"<\s*br\s*/?\s*>|</\s*p\s*>", RegexOptions.IgnoreCase)]
    private static partial Regex LineBreaks();

    [GeneratedRegex(@"<[^>]*>")]
    private static partial Regex Tags();

    public static string Strip(string value)
    {
        if (value.IndexOf('<') < 0 && value.IndexOf('&') < 0) return value;
        var text = LineBreaks().Replace(value, "\n");
        text = Tags().Replace(text, "");
        return WebUtility.HtmlDecode(text).Trim();
    }
}

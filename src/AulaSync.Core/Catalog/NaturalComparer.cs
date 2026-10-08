using System.Globalization;

namespace AulaSync.Core;

// "1A" < "2A" < "10A": talstykker sammenlignes som tal, tekststykker med dansk sortering uden store/små bogstaver.
public sealed class NaturalComparer : IComparer<string>
{
    public static readonly NaturalComparer Danish = new(CultureInfo.GetCultureInfo("da-DK"));

    readonly CompareInfo _compare;

    NaturalComparer(CultureInfo culture) => _compare = culture.CompareInfo;

    public int Compare(string? x, string? y)
    {
        if (ReferenceEquals(x, y)) return 0;
        if (x is null) return -1;
        if (y is null) return 1;

        int i = 0, j = 0;
        while (i < x.Length && j < y.Length)
        {
            if (char.IsAsciiDigit(x[i]) && char.IsAsciiDigit(y[j]))
            {
                int si = i, sj = j;
                while (i < x.Length && char.IsAsciiDigit(x[i])) i++;
                while (j < y.Length && char.IsAsciiDigit(y[j])) j++;
                var a = x[si..i].TrimStart('0');
                var b = y[sj..j].TrimStart('0');
                if (a.Length != b.Length) return a.Length.CompareTo(b.Length);
                int c = string.CompareOrdinal(a, b);
                if (c != 0) return c;
            }
            else
            {
                int si = i, sj = j;
                while (i < x.Length && !char.IsAsciiDigit(x[i])) i++;
                while (j < y.Length && !char.IsAsciiDigit(y[j])) j++;
                int c = _compare.Compare(x[si..i], y[sj..j], CompareOptions.IgnoreCase);
                if (c != 0) return c;
            }
        }
        return (x.Length - i).CompareTo(y.Length - j);
    }
}

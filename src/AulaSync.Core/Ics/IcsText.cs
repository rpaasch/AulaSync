using System.Globalization;
using System.Text;

namespace AulaSync.Core;

public static class IcsText
{
    const int MaxLineBytes = 75;

    // RFC 5545 §3.3.11: TEXT-værdier.
    public static string Escape(string value)
    {
        var sb = new StringBuilder(value.Length);
        for (int i = 0; i < value.Length; i++)
        {
            char c = value[i];
            switch (c)
            {
                case '\\': sb.Append(@"\\"); break;
                case ';': sb.Append(@"\;"); break;
                case ',': sb.Append(@"\,"); break;
                case '\r':
                    sb.Append(@"\n");
                    if (i + 1 < value.Length && value[i + 1] == '\n') i++;
                    break;
                case '\n': sb.Append(@"\n"); break;
                default: sb.Append(c); break;
            }
        }
        return sb.ToString();
    }

    // RFC 5545 §3.1: højst 75 bytes pr. fysisk linje; fortsættelseslinjer starter med mellemrum.
    // Foldes kun mellem hele tekstelementer, så UTF-8-tegn og emoji aldrig splittes.
    public static string Fold(string line)
    {
        var sb = new StringBuilder(line.Length + 8);
        int bytes = 0;
        var elements = StringInfo.GetTextElementEnumerator(line);
        while (elements.MoveNext())
        {
            var element = elements.GetTextElement();
            int size = Encoding.UTF8.GetByteCount(element);
            if (bytes + size > MaxLineBytes)
            {
                sb.Append("\r\n ");
                bytes = 1;
            }
            sb.Append(element);
            bytes += size;
        }
        return sb.Append("\r\n").ToString();
    }
}

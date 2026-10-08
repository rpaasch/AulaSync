using System.Text;
using AulaSync.Core;

namespace AulaSync.Tests;

public class IcsTextTests
{
    [Fact]
    public void Escape_handles_special_characters_and_all_newline_styles() =>
        Assert.Equal(@"a\\b\;c\,d\ne\nf\ng", IcsText.Escape("a\\b;c,d\r\ne\nf\rg"));

    [Fact]
    public void Escape_leaves_danish_letters() => Assert.Equal("Æble Øst Å", IcsText.Escape("Æble Øst Å"));

    [Fact]
    public void Short_line_is_not_folded() => Assert.Equal("ABC\r\n", IcsText.Fold("ABC"));

    [Fact]
    public void Exactly_75_bytes_is_not_folded()
    {
        var line = new string('x', 75);
        Assert.Equal(line + "\r\n", IcsText.Fold(line));
    }

    [Fact]
    public void Byte_76_moves_to_continuation_line()
    {
        var line = new string('x', 76);
        Assert.Equal(new string('x', 75) + "\r\n x\r\n", IcsText.Fold(line));
    }

    [Theory]
    [InlineData("DESCRIPTION:")]
    [InlineData("SUMMARY:x")]
    [InlineData("LOCATION:xy")]
    public void Danish_letters_are_never_split(string prefix) =>
        AssertFoldedCorrectly(prefix + string.Concat(Enumerable.Repeat("æøå", 60)));

    [Fact]
    public void Emoji_are_never_split() =>
        AssertFoldedCorrectly("SUMMARY:" + string.Concat(Enumerable.Repeat("😀", 50)));

    [Fact]
    public void Very_long_line_folds_into_many_lines() =>
        AssertFoldedCorrectly("DESCRIPTION:" + string.Concat(Enumerable.Repeat("Anna Eksempel (AE)\\, ", 40)));

    static void AssertFoldedCorrectly(string line)
    {
        var folded = IcsText.Fold(line);
        Assert.EndsWith("\r\n", folded);
        var physical = folded[..^2].Split("\r\n");
        foreach (var p in physical) Assert.True(Encoding.UTF8.GetByteCount(p) <= 75, $"Linje for lang: {Encoding.UTF8.GetByteCount(p)} bytes");
        Assert.All(physical.Skip(1), p => Assert.StartsWith(" ", p));
        Assert.Equal(line, string.Concat(physical.Select((p, i) => i == 0 ? p : p[1..])));
    }

    [Theory]
    [InlineData("<p>Dansk &amp; matematik</p>", "Dansk & matematik")]
    [InlineData("a<br>b<br/>c", "a\nb\nc")]
    [InlineData("<b>Personalemøde</b>", "Personalemøde")]
    [InlineData("Almindelig tekst", "Almindelig tekst")]
    [InlineData("", "")]
    public void StripHtml(string input, string expected) => Assert.Equal(expected, HtmlText.Strip(input));
}

using AulaSync.Core;
using static AulaSync.Tests.TestData;

namespace AulaSync.Tests;

public class AulaDatesTests
{
    [Fact]
    public void Summer_start_with_space() =>
        Assert.Equal("2026-04-13 00:00:00.0000+02:00", AulaDates.Format(new DateOnly(2026, 4, 13), false, ' ', Copenhagen));

    [Fact]
    public void Winter_end_with_T() =>
        Assert.Equal("2026-01-05T23:59:59.9990+01:00", AulaDates.Format(new DateOnly(2026, 1, 5), true, 'T', Copenhagen));

    [Fact]
    public void Dst_change_day_uses_offset_at_that_time() =>
        Assert.Equal("2026-03-29 23:59:59.9990+02:00", AulaDates.Format(new DateOnly(2026, 3, 29), true, ' ', Copenhagen));
}

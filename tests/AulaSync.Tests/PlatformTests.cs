using System.Xml.Linq;
using AulaSync.Core;
using static AulaSync.Tests.TestData;

namespace AulaSync.Tests;

public class PlatformTests
{
    [Fact]
    public void Second_instance_is_refused_until_first_is_released()
    {
        using var dir = new TempDir();
        var path = dir.File("aulasync.lock");
        using (var first = SingleInstance.TryAcquire(path))
        {
            Assert.NotNull(first);
            Assert.Null(SingleInstance.TryAcquire(path));
        }
        using var again = SingleInstance.TryAcquire(path);
        Assert.NotNull(again);
    }

    [Fact]
    public async Task Signal_reaches_the_running_instance()
    {
        var name = "AulaSync-test-" + Guid.NewGuid().ToString("N")[..8];
        int shown = 0;
        using var channel = new InstanceChannel(name);
        channel.Listen(() => Interlocked.Increment(ref shown));

        Assert.True(await InstanceChannel.SignalShowAsync(name, TimeSpan.FromSeconds(2)));
        await WaitUntil(() => shown == 1);
        Assert.True(await InstanceChannel.SignalShowAsync(name, TimeSpan.FromSeconds(2)));
        await WaitUntil(() => shown == 2);
    }

    // Startes AulaSync igen, mens den kørende instans stadig behandler en tidligere besked, må beskeden ikke gå tabt.
    // Før lukkede lytteren efter behandlingen, og en forbindelse, der ventede i dens kø, forsvandt med den.
    [Fact]
    public async Task Signal_while_the_previous_one_is_handled_is_not_lost()
    {
        var name = "AulaSync-test-" + Guid.NewGuid().ToString("N")[..8];
        using var release = new SemaphoreSlim(0);
        int shown = 0;
        using var channel = new InstanceChannel(name);
        channel.Listen(() => { if (Interlocked.Increment(ref shown) == 1) release.Wait(TimeSpan.FromSeconds(5)); });

        Assert.True(await InstanceChannel.SignalShowAsync(name, TimeSpan.FromSeconds(2)));
        await WaitUntil(() => shown == 1);
        Assert.True(await InstanceChannel.SignalShowAsync(name, TimeSpan.FromSeconds(2))); // mens den første behandles
        release.Release();
        await WaitUntil(() => shown == 2);
    }

    [Fact]
    public async Task Signal_without_running_instance_returns_false() =>
        Assert.False(await InstanceChannel.SignalShowAsync("AulaSync-ingen-" + Guid.NewGuid().ToString("N")[..8], TimeSpan.FromMilliseconds(300)));

    [Fact]
    public void LaunchAgent_plist_runs_app_silently_at_login()
    {
        var xml = XDocument.Parse(LaunchAgent.CreatePlist("/Applications/AulaSync & co.app/Contents/MacOS/AulaSync"));
        var values = xml.Descendants("dict").First().Elements().ToList();
        Assert.Contains(values, e => e.Value == "dk.rpaasch.aulasync");
        var args = xml.Descendants("array").Single().Elements("string").Select(e => e.Value).ToList();
        Assert.Equal(["/Applications/AulaSync & co.app/Contents/MacOS/AulaSync", "--silent"], args);
        Assert.Contains(values, e => e.Name == "true");
    }

    [Fact]
    public void LaunchAgent_path() =>
        Assert.Equal(Path.Combine("/Users/x", "Library", "LaunchAgents", "dk.rpaasch.aulasync.plist"), LaunchAgent.PlistPath("/Users/x"));
}

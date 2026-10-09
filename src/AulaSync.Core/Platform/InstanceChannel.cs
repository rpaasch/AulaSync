using System.IO.Pipes;
using System.Text;

namespace AulaSync.Core;

// Navngivet pipe (på Mac en Unix-socket bag kulisserne): en anden start beder den kørende instans om at vise hovedvinduet,
// og installationsprogrammet (AulaSync.exe --quit) beder den om at afslutte.
public sealed class InstanceChannel(string name) : IDisposable
{
    const string ShowCommand = "show";
    const string QuitCommand = "quit";
    readonly CancellationTokenSource _cts = new();

    public static string DefaultName => $"AulaSync-{Environment.UserName}";

    public void Listen(Action onShow, Action? onQuit = null) => _ = Task.Run(() => LoopAsync(onShow, onQuit ?? (() => { })));

    // Den næste lytter oprettes, før den nuværende lukkes. Ellers forsvinder en ny start, der kobler på, mens en
    // besked behandles, med den lukkede lytter (Unix), eller den finder ingen ledig lytter (Windows).
    async Task LoopAsync(Action onShow, Action onQuit)
    {
        NamedPipeServerStream? next = null;
        try
        {
            while (!_cts.IsCancellationRequested)
            {
                try
                {
                    await using var server = next ?? CreateServer();
                    next = null;
                    await server.WaitForConnectionAsync(_cts.Token);
                    next = CreateServer();
                    using var reader = new StreamReader(server, Encoding.UTF8);
                    switch (await reader.ReadLineAsync(_cts.Token))
                    {
                        case ShowCommand: onShow(); break;
                        case QuitCommand: onQuit(); break;
                    }
                }
                catch (OperationCanceledException) { return; }
                catch (IOException) { await Task.Delay(200); }
            }
        }
        finally { next?.Dispose(); }
    }

    // To instanser: den, der behandler en besked, og den næste, der allerede lytter.
    NamedPipeServerStream CreateServer() => new(name, PipeDirection.In, 2, PipeTransmissionMode.Byte,
        PipeOptions.Asynchronous | PipeOptions.CurrentUserOnly);

    public static Task<bool> SignalShowAsync(string name, TimeSpan timeout) => SignalAsync(name, ShowCommand, timeout);

    public static Task<bool> SignalQuitAsync(string name, TimeSpan timeout) => SignalAsync(name, QuitCommand, timeout);

    static async Task<bool> SignalAsync(string name, string command, TimeSpan timeout)
    {
        try
        {
            // Windows: et installationsprogram, der kører som administrator, skal kunne nå en AulaSync, der ikke gør. Med
            // CurrentUserOnly sammenligner klienten pipens ejer (brugeren) med sin egen ejer (Administratorer) og afviser.
            // Serverens adgangsliste giver stadig kun brugeren selv adgang, og klienten sender kun "show" eller "quit".
            var options = OperatingSystem.IsWindows() ? PipeOptions.None : PipeOptions.CurrentUserOnly;
            await using var client = new NamedPipeClientStream(".", name, PipeDirection.Out, options);
            await client.ConnectAsync((int)timeout.TotalMilliseconds);
            await client.WriteAsync(Encoding.UTF8.GetBytes(command + "\n"));
            await client.FlushAsync();
            return true;
        }
        catch (Exception ex) when (ex is TimeoutException or IOException or UnauthorizedAccessException) { return false; }
    }

    public void Dispose() => _cts.Cancel();
}

using System.IO.Pipes;
using System.Text;

namespace AulaSync.Core;

// Navngivet pipe (på Mac en Unix-socket bag kulisserne): en anden start beder den kørende instans om at vise hovedvinduet.
public sealed class InstanceChannel(string name) : IDisposable
{
    const string ShowCommand = "show";
    readonly CancellationTokenSource _cts = new();

    public static string DefaultName => $"AulaSync-{Environment.UserName}";

    public void Listen(Action onShow) => _ = Task.Run(() => LoopAsync(onShow));

    // Den næste lytter oprettes, før den nuværende lukkes. Ellers forsvinder en ny start, der kobler på, mens en
    // besked behandles, med den lukkede lytter (Unix), eller den finder ingen ledig lytter (Windows).
    async Task LoopAsync(Action onShow)
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
                    if (await reader.ReadLineAsync(_cts.Token) == ShowCommand) onShow();
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

    public static async Task<bool> SignalShowAsync(string name, TimeSpan timeout)
    {
        try
        {
            await using var client = new NamedPipeClientStream(".", name, PipeDirection.Out, PipeOptions.CurrentUserOnly);
            await client.ConnectAsync((int)timeout.TotalMilliseconds);
            await client.WriteAsync(Encoding.UTF8.GetBytes(ShowCommand + "\n"));
            await client.FlushAsync();
            return true;
        }
        catch (Exception ex) when (ex is TimeoutException or IOException or UnauthorizedAccessException) { return false; }
    }

    public void Dispose() => _cts.Cancel();
}

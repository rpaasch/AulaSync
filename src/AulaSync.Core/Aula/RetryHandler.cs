namespace AulaSync.Core;

// Prøver igen ved 5xx og netværksfejl: 3 forsøg med 1 s og 2 s pause.
public sealed class RetryHandler(HttpMessageHandler inner, Func<TimeSpan, CancellationToken, Task>? delay = null, int attempts = 3)
    : DelegatingHandler(inner)
{
    readonly Func<TimeSpan, CancellationToken, Task> _delay = delay ?? Task.Delay;

    protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken ct)
    {
        if (request.Content is not null) await request.Content.LoadIntoBufferAsync(ct);
        for (int attempt = 1; ; attempt++)
        {
            try
            {
                var response = await base.SendAsync(request, ct);
                if ((int)response.StatusCode < 500 || attempt >= attempts) return response;
                response.Dispose();
            }
            catch (HttpRequestException) when (attempt < attempts) { }
            catch (TaskCanceledException) when (attempt < attempts && !ct.IsCancellationRequested) { }
            await _delay(TimeSpan.FromSeconds(attempt), ct);
        }
    }
}

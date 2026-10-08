using System.Net;

namespace AulaSync.Core;

public sealed class AulaSession
{
    public const string SessionCookie = "PHPSESSID";
    public const string CsrfCookie = "Csrfp-Token";
    static readonly Uri AulaUri = new("https://www.aula.dk/");

    readonly CookieContainer _cookies = new();

    public AulaSession(IEnumerable<Cookie> cookies)
    {
        var list = cookies.ToList();
        if (!HasRequiredCookies(list)) throw new ArgumentException("Login-cookies fra Aula mangler", nameof(cookies));
        foreach (var c in list)
        {
            var domain = string.IsNullOrEmpty(c.Domain) ? AulaUri.Host : c.Domain;
            _cookies.Add(new Cookie(c.Name, c.Value, string.IsNullOrEmpty(c.Path) ? "/" : c.Path, domain));
        }
    }

    public static bool HasRequiredCookies(IEnumerable<Cookie> cookies)
    {
        var list = cookies as ICollection<Cookie> ?? cookies.ToList();
        return list.Any(c => c.Name == SessionCookie && c.Value != "") && list.Any(c => c.Name == CsrfCookie && c.Value != "");
    }

    public HttpClient CreateHttpClient(Func<TimeSpan, CancellationToken, Task>? delay = null, HttpMessageHandler? inner = null)
    {
        var transport = inner ?? new HttpClientHandler { CookieContainer = _cookies, UseCookies = true };
        var http = new HttpClient(new RetryHandler(transport, delay)) { Timeout = TimeSpan.FromSeconds(30) };
        http.DefaultRequestHeaders.Accept.ParseAdd("application/json, text/plain, */*");
        var csrf = _cookies.GetCookies(AulaUri).FirstOrDefault(c => c.Name == CsrfCookie);
        if (csrf is not null) http.DefaultRequestHeaders.Add("csrfp-token", csrf.Value);
        return http;
    }
}

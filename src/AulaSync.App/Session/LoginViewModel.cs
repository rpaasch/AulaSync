using System.Net;
using AulaSync.Core;
using CommunityToolkit.Mvvm.ComponentModel;

namespace AulaSync.App;

public sealed partial class LoginViewModel(SessionController session) : ObservableObject
{
    [ObservableProperty] public partial string? Error { get; set; }

    [ObservableProperty] public partial bool IsBusy { get; set; }

    public event Action<Profile>? SignedIn;

    // Kaldes af LoginView, når Aula har sat begge login-cookies. Returnerer false, hvis login skal prøves igen.
    public async Task<bool> CookiesFoundAsync(IReadOnlyList<Cookie> cookies)
    {
        IsBusy = true;
        Error = null;
        try
        {
            // Ikke SignedIn?.Invoke(await …): med ?. springes hele kaldet — og dermed login — over, når ingen lytter.
            var profile = await session.SignInAsync(cookies, CancellationToken.None);
            SignedIn?.Invoke(profile);
            return true;
        }
        catch (SessionExpiredException) { Error = "Aula afviste login. Prøv igen."; }
        catch (Exception ex) when (ex is AulaException or HttpRequestException or TaskCanceledException)
        {
            Error = $"Du er logget ind, men AulaSync kunne ikke hente din profil: {ex.Message}";
        }
        finally { IsBusy = false; }
        return false;
    }
}

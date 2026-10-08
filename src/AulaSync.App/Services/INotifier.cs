namespace AulaSync.App;

// En kort besked uden for vinduerne (spec §5). I dag AulaSyncs egen boks (NotificationBox). I den pakkede app kan en
// systemnotifikation (Notifikationscenter / Windows-toast) gøre det samme bag denne grænseflade (Task 20–21).
public interface INotifier
{
    // Brugeren klikkede på en besked uden egen handling (genlogin: login-vinduet).
    event Action? Clicked;

    // Brugeren lukkede beskeden (✕).
    event Action? Dismissed;

    // onClick: hvad et klik på netop denne besked gør (fx åbn hovedvinduet); uden den giver klikket Clicked.
    void Show(string text, Action? onClick = null);
}

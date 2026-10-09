namespace AulaSync.App;

// AulaSync 2 (Windows) kørte med mutexen "AulaSync_SingleInstance" og en kalender-server på port 9876. Kører den stadig efter
// opgraderingen (den startede ved login), siger AulaSync 3 til, så man ikke har to ikoner og to programmer, der svarer på
// kalenderadresser. AulaSync 3 overtager startpunktet ved login (samme Run-værdi "AulaSync"), når den starter eller
// installeres (IAutostart.Retarget, installationsprogrammet), og ellers ved første start.
public static class OldVersion
{
    public const string MutexName = "AulaSync_SingleInstance";
    public const string Text = "Den gamle AulaSync (version 2) kører stadig. Højreklik på dens ikon i systembakken, og vælg Afslut.";

    public static bool IsRunning(string mutexName = MutexName)
    {
        if (!Mutex.TryOpenExisting(mutexName, out var mutex)) return false;
        mutex.Dispose();
        return true;
    }
}

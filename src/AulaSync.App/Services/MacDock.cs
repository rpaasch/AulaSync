using System.Runtime.InteropServices;

namespace AulaSync.App;

// Mac: Dock-ikon, mens et vindue er åbent; ellers kun menulinjen (spec §3.2, spiken punkt 6a). Desuden det, notifikations-
// boksen skal bruge fra AppKit.
public static class MacDock
{
    [DllImport("/usr/lib/libobjc.A.dylib")] static extern IntPtr objc_getClass(string name);
    [DllImport("/usr/lib/libobjc.A.dylib")] static extern IntPtr sel_registerName(string name);
    [DllImport("/usr/lib/libobjc.A.dylib", EntryPoint = "objc_msgSend")] static extern IntPtr Send(IntPtr receiver, IntPtr selector);
    [DllImport("/usr/lib/libobjc.A.dylib", EntryPoint = "objc_msgSend")] static extern void SendLong(IntPtr receiver, IntPtr selector, long value);
    [DllImport("/usr/lib/libobjc.A.dylib", EntryPoint = "objc_msgSend")] static extern void SendBool(IntPtr receiver, IntPtr selector, [MarshalAs(UnmanagedType.I1)] bool value);
    [DllImport("/usr/lib/libobjc.A.dylib", EntryPoint = "objc_msgSend")] static extern void SendPointer(IntPtr receiver, IntPtr selector, IntPtr value);
    [DllImport("/usr/lib/libobjc.A.dylib", EntryPoint = "objc_msgSend")] [return: MarshalAs(UnmanagedType.I1)] static extern bool GetBool(IntPtr receiver, IntPtr selector);

    static IntPtr App => Send(objc_getClass("NSApplication"), sel_registerName("sharedApplication"));

    // NSApplicationActivationPolicyRegular = 0 (Dock-ikon), Accessory = 1 (kun menulinje).
    public static void SetVisible(bool visible)
    {
        if (!OperatingSystem.IsMacOS()) return;
        var app = App;
        SendLong(app, sel_registerName("setActivationPolicy:"), visible ? 0 : 1);
        if (visible) SendBool(app, sel_registerName("activateIgnoringOtherApps:"), true);
    }

    // Et klik i notifikationsboksen aktiverer AulaSync. Lukkes boksen med ✕, giver [NSApp hide:] fokus tilbage til det
    // program, der var aktivt før (macOS gør det ikke selv, når et vindue lukkes).
    public static void HideApp()
    {
        if (OperatingSystem.IsMacOS()) SendPointer(App, sel_registerName("hide:"), IntPtr.Zero);
    }

    // App-menuens "Skjul andre" og "Vis alle".
    public static void HideOthers()
    {
        if (OperatingSystem.IsMacOS()) SendPointer(App, sel_registerName("hideOtherApplications:"), IntPtr.Zero);
    }

    public static void ShowAll()
    {
        if (OperatingSystem.IsMacOS()) SendPointer(App, sel_registerName("unhideAllApplications:"), IntPtr.Zero);
    }

    // Modparten: en skjult app kan ikke vise boksen, så den vises igen uden at blive aktiv.
    public static void UnhideWithoutActivation()
    {
        if (!OperatingSystem.IsMacOS()) return;
        var app = App;
        if (GetBool(app, sel_registerName("isHidden"))) Send(app, sel_registerName("unhideWithoutActivation"));
    }

    // NSWindowCollectionBehaviorCanJoinAllSpaces (1) | FullScreenAuxiliary (256): vinduet vises i alle Spaces og over
    // programmer i fuld skærm.
    public static void ShowOnAllSpaces(IntPtr nsWindow)
    {
        if (OperatingSystem.IsMacOS() && nsWindow != IntPtr.Zero) SendLong(nsWindow, sel_registerName("setCollectionBehavior:"), 1 | 256);
    }
}

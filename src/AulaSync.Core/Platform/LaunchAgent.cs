using System.Security;

namespace AulaSync.Core;

public static class LaunchAgent
{
    public const string Label = "dk.rpaasch.aulasync";

    public static string PlistPath(string home) => Path.Combine(home, "Library", "LaunchAgents", Label + ".plist");

    public static string CreatePlist(string executablePath) => $"""
        <?xml version="1.0" encoding="UTF-8"?>
        <!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
        <plist version="1.0">
        <dict>
            <key>Label</key>
            <string>{Label}</string>
            <key>ProgramArguments</key>
            <array>
                <string>{SecurityElement.Escape(executablePath)}</string>
                <string>--silent</string>
            </array>
            <key>RunAtLoad</key>
            <true/>
        </dict>
        </plist>
        """;
}

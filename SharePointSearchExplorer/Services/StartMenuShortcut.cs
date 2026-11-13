using System.Diagnostics;

namespace SharePointSearchExplorer.Services;

/// <summary>
/// A Start menu shortcut for the app's .exe, for the current user only (no admin rights or installer needed).
/// </summary>
public static class StartMenuShortcut
{
    private const string Name = "SharePoint Search Explorer.lnk";

    /// <summary>%APPDATA%\Microsoft\Windows\Start Menu\Programs\SharePoint Search Explorer.lnk</summary>
    public static string ShortcutPath { get; } =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Programs), Name);

    public static bool Exists => File.Exists(ShortcutPath);

    /// <summary>Creates (or replaces) the shortcut, pointing at the running .exe.</summary>
    public static void Create()
    {
        var exe = UpdateService.ExecutablePath ?? throw new InvalidOperationException("Couldn't find where the app is running from.");
        Directory.CreateDirectory(Path.GetDirectoryName(ShortcutPath)!);

        // WScript.Shell writes .lnk files without hand-written IShellLink interop.
        var shellType = Type.GetTypeFromProgID("WScript.Shell") ?? throw new InvalidOperationException("Windows Script Host isn't available.");
        dynamic shell = Activator.CreateInstance(shellType)!;
        try
        {
            dynamic link = shell.CreateShortcut(ShortcutPath);
            link.TargetPath = exe;
            link.WorkingDirectory = Path.GetDirectoryName(exe);
            link.IconLocation = exe + ",0";
            link.Description = "Explore SharePoint search results and managed properties.";
            link.Save();
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(shell);
        }
    }

    public static void Remove()
    {
        if (File.Exists(ShortcutPath)) File.Delete(ShortcutPath);
    }

    /// <summary>
    /// At startup: if there's a shortcut but it points somewhere else, points it at this .exe, so moving the app to
    /// another folder doesn't leave the shortcut opening nothing.
    /// </summary>
    public static void KeepCurrent()
    {
        if (!Exists || UpdateService.ExecutablePath is not { } exe) return;
        _ = Task.Run(() =>
        {
            try
            {
                if (!string.Equals(ReadTarget(), exe, StringComparison.OrdinalIgnoreCase)) Create();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Could not update the Start menu shortcut: {ex.Message}");
            }
        });
    }

    private static string? ReadTarget()
    {
        var shellType = Type.GetTypeFromProgID("WScript.Shell");
        if (shellType == null) return null;
        dynamic shell = Activator.CreateInstance(shellType)!;
        try
        {
            return (string)shell.CreateShortcut(ShortcutPath).TargetPath;
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(shell);
        }
    }
}

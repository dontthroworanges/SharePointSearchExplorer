using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Windows.Graphics;

namespace SharePointSearchExplorer.Views;

/// <summary>
/// One Server Response tab in its own window. It keeps a snapshot of the text, so new searches don't change
/// it and several searches can be compared side by side.
/// </summary>
public sealed partial class CodeWindow : Window
{
    private static readonly List<CodeWindow> s_open = new();

    private CodeWindow(string tabName, string request, string text)
    {
        InitializeComponent();

        Title = tabName;
        ExtendsContentIntoTitleBar = true;
        SetTitleBar(WindowTitleBar);
        WindowTitleBar.Title = tabName;
        AppWindow.SetIcon(Path.Combine(AppContext.BaseDirectory, "Assets", "AppIcon.ico"));
        RequestText.Text = request;
        Code.Text = text;
        Finder.Target = Code;

        ApplyTheme(App.Settings.Data.Theme);
        Closed += (_, _) => s_open.Remove(this);
    }

    /// <summary>Opens a new window for the given text, cascaded over the main window.</summary>
    public static void Show(string tabName, string request, string text)
    {
        var window = new CodeWindow(tabName, request, text);
        var main = App.MainWindow.AppWindow;
        int offset = 32 * (s_open.Count % 8);
        window.AppWindow.MoveAndResize(new RectInt32(main.Position.X + 80 + offset, main.Position.Y + 60 + offset, 900, 700));
        if (window.AppWindow.Presenter is OverlappedPresenter presenter)
        {
            presenter.PreferredMinimumWidth = 480;
            presenter.PreferredMinimumHeight = 320;
        }
        s_open.Add(window);
        window.Activate();
    }

    /// <summary>Follows the app theme set in Settings › Appearance.</summary>
    public static void ApplyThemeToAll(string theme)
    {
        foreach (var window in s_open) window.ApplyTheme(theme);
    }

    /// <summary>Picks up a new accent colour set in Settings › Appearance.</summary>
    public static void RefreshAccentColorForAll()
    {
        foreach (var window in s_open) Helpers.AccentColors.Refresh(window.RootGrid);
    }

    /// <summary>Pop-outs belong to the main window; close them with it.</summary>
    public static void CloseAll()
    {
        foreach (var window in s_open.ToList()) window.Close();
    }

    private void ApplyTheme(string theme)
    {
        RootGrid.RequestedTheme = theme switch
        {
            "Light" => ElementTheme.Light,
            "Dark" => ElementTheme.Dark,
            _ => ElementTheme.Default,
        };
        AppWindow.TitleBar.PreferredTheme = theme switch
        {
            "Light" => TitleBarTheme.Light,
            "Dark" => TitleBarTheme.Dark,
            _ => TitleBarTheme.UseDefaultAppMode,
        };
    }

    private void Copy_Click(object sender, RoutedEventArgs e) => PropertyClipboard.CopyText(Code.Text);
}

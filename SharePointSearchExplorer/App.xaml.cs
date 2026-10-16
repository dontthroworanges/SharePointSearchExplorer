using System.Diagnostics;
using Microsoft.UI.Xaml;
using SharePointSearchExplorer.Services;
using SharePointSearchExplorer.ViewModels;

namespace SharePointSearchExplorer;

public partial class App : Application
{
    public static SettingsStore Settings { get; private set; } = null!;
    public static AuthService Auth { get; private set; } = null!;
    public static SearchService SearchService { get; private set; } = null!;
    public static SearchViewModel Search { get; private set; } = null!;
    public static MainWindow MainWindow { get; private set; } = null!;

    /// <summary>Native window handle, used to parent the MSAL browser sign-in.</summary>
    public static IntPtr WindowHandle { get; private set; }

    public App()
    {
        InitializeComponent();
        UnhandledException += (_, e) =>
        {
            Debug.WriteLine($"Unhandled exception: {e.Exception}");            e.Handled = true;
        };
    }

    protected override void OnLaunched(LaunchActivatedEventArgs args)
    {
        Settings = SettingsStore.Load();
        Helpers.DisplayOptions.Current.WordWrap = Settings.Data.WordWrap;
        Helpers.DisplayOptions.Current.WrapCodeBlocks = Settings.Data.WrapCodeBlocks;
        Helpers.DisplayOptions.Current.UseFileTypeIcons = Settings.Data.UseFileTypeIcons;
        Helpers.DisplayOptions.Current.DenseResults = Settings.Data.ResultDensity == "Dense";
        Helpers.AccentColors.Apply(Settings.Data.AccentColor);
        Auth = new AuthService();
        SearchService = new SearchService(Auth);
        Search = new SearchViewModel(Settings, SearchService, Auth);

        MainWindow = new MainWindow();
        WindowHandle = WinRT.Interop.WindowNative.GetWindowHandle(MainWindow);
        MainWindow.Activate();
    }
}

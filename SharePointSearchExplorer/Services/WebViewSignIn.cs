using Microsoft.Identity.Client.Extensibility;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.Web.WebView2.Core;
using Windows.Graphics;

namespace SharePointSearchExplorer.Services;

public sealed class SignInCanceledException() : OperationCanceledException("Sign-in was canceled.");

/// <summary>
/// Embedded Microsoft sign-in page (WinUI WebView2) used by MSAL as its web UI.
///
/// This is what the original tool's "SharePoint Online" method does: the built-in
/// SharePoint Online Management Shell client redirects to https://oauth.spops.microsoft.com/,
/// which only an embedded browser can intercept. Cookies are kept in a persistent
/// WebView2 profile, and Windows single sign-on is allowed, so an existing session
/// usually signs in without asking for a password.
/// </summary>
public sealed class WebViewSignIn : ICustomWebUi
{
    private static string UserDataFolder => Path.Combine(SettingsStore.DataFolder, "WebView2");

    /// <summary>Deletes the sign-in window's cookies and cache (best effort: WebView2 may still hold some files).</summary>
    public static void ClearBrowsingData()
    {
        try
        {
            if (Directory.Exists(UserDataFolder)) Directory.Delete(UserDataFolder, recursive: true);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            System.Diagnostics.Debug.WriteLine($"Could not clear sign-in data: {ex.Message}");
        }
    }

    public Task<Uri> AcquireAuthorizationCodeAsync(Uri authorizationUri, Uri redirectUri, CancellationToken cancellationToken)
    {
        var tcs = new TaskCompletionSource<Uri>(TaskCreationOptions.RunContinuationsAsynchronously);
        var dispatcher = App.MainWindow.DispatcherQueue;
        if (!dispatcher.TryEnqueue(DispatcherQueuePriority.Normal, () => ShowWindow(authorizationUri, redirectUri, tcs, cancellationToken)))
            tcs.TrySetException(new InvalidOperationException("Could not open the sign-in window."));
        return tcs.Task;
    }

    private static async void ShowWindow(Uri authorizationUri, Uri redirectUri, TaskCompletionSource<Uri> tcs, CancellationToken ct)
    {
        var webView = new WebView2();
        var progress = new ProgressRing { IsActive = true, Width = 40, Height = 40 };
        var root = new Grid { Children = { webView, progress } };
        var window = new Window { Content = root, Title = "Sign in to SharePoint Online" };

        root.RequestedTheme = App.MainWindow.Content is FrameworkElement main ? main.ActualTheme : ElementTheme.Default;
        window.AppWindow.SetIcon(Path.Combine(AppContext.BaseDirectory, "Assets", "AppIcon.ico"));
        CenterOverMainWindow(window.AppWindow, new SizeInt32(520, 720));
        if (window.AppWindow.Presenter is OverlappedPresenter presenter)
        {
            presenter.IsMinimizable = false;
            presenter.IsMaximizable = false;
        }

        bool completed = false;
        void Finish(Action complete)
        {
            if (completed) return;
            completed = true;
            complete();
            window.Close();
        }

        var registration = ct.Register(() => window.DispatcherQueue.TryEnqueue(() => Finish(() => tcs.TrySetCanceled(ct))));
        window.Closed += (_, _) =>
        {
            registration.Dispose();
            if (!completed)
            {
                completed = true;
                tcs.TrySetException(new SignInCanceledException());
            }
        };

        window.Activate();

        try
        {
            Directory.CreateDirectory(UserDataFolder);
            var options = new CoreWebView2EnvironmentOptions { AllowSingleSignOnUsingOSPrimaryAccount = true };
            var environment = await CoreWebView2Environment.CreateWithOptionsAsync("", UserDataFolder, options);
            await webView.EnsureCoreWebView2Async(environment);
        }
        catch (Exception ex)
        {
            Finish(() => tcs.TrySetException(new InvalidOperationException(
                "The sign-in window needs the Microsoft Edge WebView2 Runtime. " + ex.Message, ex)));
            return;
        }

        var core = webView.CoreWebView2;
        core.Settings.AreDevToolsEnabled = false;
        core.Settings.IsStatusBarEnabled = false;
        core.NavigationStarting += (_, e) =>
        {
            // The redirect carries the authorization code (or an error) back to MSAL.
            if (e.Uri.StartsWith(redirectUri.GetLeftPart(UriPartial.Path), StringComparison.OrdinalIgnoreCase))
            {
                e.Cancel = true;
                var result = new Uri(e.Uri);
                Finish(() => tcs.TrySetResult(result));
            }
        };
        core.NavigationCompleted += (_, _) => progress.IsActive = false;
        core.Navigate(authorizationUri.AbsoluteUri);
    }

    private static void CenterOverMainWindow(AppWindow appWindow, SizeInt32 size)
    {
        var main = App.MainWindow.AppWindow;
        int x = main.Position.X + (main.Size.Width - size.Width) / 2;
        int y = main.Position.Y + (main.Size.Height - size.Height) / 2;
        appWindow.MoveAndResize(new RectInt32(Math.Max(0, x), Math.Max(0, y), size.Width, size.Height));
    }
}

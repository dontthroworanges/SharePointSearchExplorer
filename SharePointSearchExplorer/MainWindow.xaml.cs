using System.ComponentModel;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Automation;
using Microsoft.UI.Xaml.Automation.Peers;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using SharePointSearchExplorer.Helpers;
using SharePointSearchExplorer.Models;
using SharePointSearchExplorer.Services;
using SharePointSearchExplorer.ViewModels;
using SharePointSearchExplorer.Views;
using Windows.Graphics;

namespace SharePointSearchExplorer;

public sealed partial class MainWindow : Window
{
    private bool _dialogOpen;
    private ConnectionProfile? _watchedProfile;

    public MainWindow()
    {
        InitializeComponent();

        ExtendsContentIntoTitleBar = true;
        SetTitleBar(AppTitleBar);
        AppWindow.SetIcon(Path.Combine(AppContext.BaseDirectory, "Assets", "AppIcon.ico"));
        AppWindow.Resize(new SizeInt32(1280, 820));
        if (AppWindow.Presenter is OverlappedPresenter presenter)
        {
            presenter.PreferredMinimumWidth = 820;
            presenter.PreferredMinimumHeight = 520;
        }

        ApplyTheme(App.Settings.Data.Theme);
        // While following Windows, pick up an accent colour changed in Windows Settings.
        AccentColors.SystemAccentChanged += (_, _) => DispatcherQueue.TryEnqueue(() =>
        {
            if (AccentColors.UsesSystemAccent) ApplyAccentColor("");
        });
        // History is most-recent first, so its first entry is the last query run before the app closed.
        // (Requests saved from the Advanced Query Rules page are not query texts.)
        if (App.Settings.Data.RestoreLastQuery && App.Settings.Data.History.FirstOrDefault(h => !AdvancedRequest.IsSaved(h)) is { } lastQuery)
            Search.QueryText = lastQuery;
        ApplyAdvancedPagesVisibility();
        NavView.SelectedItem = SearchNavItem;
        ContentFrame.Navigate(typeof(SearchPage));

        App.Settings.CurrentProfileChanged += (_, _) => UpdateProfileDisplay();
        Search.PropertyChanged += Search_PropertyChanged;
        Search.ProfileRequired += async (_, _) => await EditProfileAsync(null);
        RootGrid.Loaded += RootGrid_Loaded;
        QueryBox.Loaded += (_, _) => HookQueryBoxClearButton();
        RootGrid.SizeChanged += (_, e) =>
            // TitleBar sizes its content to fit, so give the query bar an explicit share of the window width
            // (after the icon and app name, the caption buttons and drag space).
            QueryBar.Width = Math.Clamp(e.NewSize.Width - 520, 280, 820);
        Closed += (_, _) =>
        {
            PropertiesWindow.CloseAll();
            CodeWindow.CloseAll();
            App.Settings.Save();
        };

        UpdateProfileDisplay();
    }

    public SearchViewModel Search => App.Search;

    private async void RootGrid_Loaded(object sender, RoutedEventArgs e)
    {
        if (App.Settings.CurrentProfile == null)
        {
            // First run: ask for a site to connect to.
            await EditProfileAsync(null);
        }
        else
        {
            await Search.RefreshSignInStatusAsync();
        }
        QueryBox.Focus(FocusState.Programmatic);
    }

    #region Theme

    public void ApplyTheme(string theme)
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
        PropertiesWindow.ApplyThemeToAll(theme);
        CodeWindow.ApplyThemeToAll(theme);
    }

    /// <summary>Applies an accent colour ("#RRGGBB", or empty for the Windows accent colour) to every open window.</summary>
    public void ApplyAccentColor(string hex)
    {
        AccentColors.Apply(hex);
        AccentColors.Refresh(RootGrid);
        PropertiesWindow.RefreshAccentColorForAll();
        CodeWindow.RefreshAccentColorForAll();
    }

    #endregion

    #region Query box

    private async void SearchButton_Click(object sender, RoutedEventArgs e)
    {
        if (Search.IsBusy)
        {
            Search.CancelSearch();
            return;
        }
        await RunSearchAsync();
    }

    private async void QueryBox_QuerySubmitted(AutoSuggestBox sender, AutoSuggestBoxQuerySubmittedEventArgs args)
    {
        if (args.ChosenSuggestion is string chosen) Search.QueryText = chosen;
        await RunSearchAsync();
    }

    private void QueryBox_TextChanged(AutoSuggestBox sender, AutoSuggestBoxTextChangedEventArgs args)
    {
        if (args.Reason == AutoSuggestionBoxTextChangeReason.UserInput)
            sender.ItemsSource = Search.GetHistorySuggestions(sender.Text).ToList();
    }

    private void QueryBox_GotFocus(object sender, RoutedEventArgs e)
    {
        // Show recent queries when the user clicks into an empty box (not on programmatic focus at startup).
        if (App.Settings.Data.ShowRecentQueriesOnFocus && string.IsNullOrEmpty(QueryBox.Text) && QueryBox.FocusState != FocusState.Programmatic)
        {
            QueryBox.ItemsSource = Search.GetHistorySuggestions("").ToList();
            QueryBox.IsSuggestionListOpen = QueryBox.Items.Count > 0;
        }
    }

    private Button? _queryClearButton;

    /// <summary>
    /// The X in the query box clears the text; it also clears Primary Results. Only the X does this, so
    /// deleting the text by typing leaves the results in place. The X is the "DeleteButton" part of the
    /// AutoSuggestBox's inner TextBox template.
    /// </summary>
    private void HookQueryBoxClearButton()
    {
        if (_queryClearButton != null || FindDescendant(QueryBox, "DeleteButton") is not Button button) return;
        _queryClearButton = button;
        button.Click += (_, _) =>
        {
            QueryBox.IsSuggestionListOpen = false;
            Search.ClearResults();
        };
    }

    #region Removing suggestions

    private void Suggestion_PointerEntered(object sender, PointerRoutedEventArgs e) =>
        ShowRemoveButton(sender, App.Settings.Data.AllowQueryBoxDeletion);

    private void Suggestion_PointerExited(object sender, PointerRoutedEventArgs e) => ShowRemoveButton(sender, false);

    /// <summary>The X is always in the row (so the text doesn't reflow), just invisible and unclickable until hover.</summary>
    private static void ShowRemoveButton(object row, bool show)
    {
        if (row is not Grid grid || grid.Children.OfType<Button>().FirstOrDefault() is not { } button) return;
        button.Opacity = show ? 1 : 0;
        button.IsHitTestVisible = show;
    }

    private async void RemoveSuggestion_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { Tag: string query } || !App.Settings.Data.AllowQueryBoxDeletion) return;

        if (App.Settings.Data.ConfirmQueryBoxDeletion)
        {
            QueryBox.IsSuggestionListOpen = false;
            bool confirmed = await ConfirmAsync(
                "Remove query from history?",
                $"“{query}”\n\nIt will no longer be suggested in the query box or listed on the Query History page.",
                "Remove");
            if (!confirmed)
            {
                ReopenSuggestions();
                return;
            }
        }

        App.Settings.RemoveHistory(query);
        ReopenSuggestions();
    }

    /// <summary>Shows the (updated) suggestions for the current text again, with focus back in the query box.</summary>
    private void ReopenSuggestions()
    {
        QueryBox.Focus(FocusState.Programmatic);
        var suggestions = Search.GetHistorySuggestions(QueryBox.Text).ToList();
        QueryBox.ItemsSource = suggestions;
        QueryBox.IsSuggestionListOpen = suggestions.Count > 0;
    }

    #endregion

    /// <summary>
    /// A Cancel / <paramref name="primaryText"/> dialog over the main window; true when the user confirms.
    /// False without asking when another dialog is already open (WinUI allows only one at a time).
    /// </summary>
    public async Task<bool> ConfirmAsync(string title, string message, string primaryText)
    {
        if (_dialogOpen) return false;
        _dialogOpen = true;
        try
        {
            var dialog = new ContentDialog
            {
                XamlRoot = RootGrid.XamlRoot,
                RequestedTheme = RootGrid.ActualTheme,
                Title = title,
                Content = new TextBlock { Text = message, TextWrapping = TextWrapping.Wrap },
                PrimaryButtonText = primaryText,
                CloseButtonText = "Cancel",
                DefaultButton = ContentDialogButton.Close,
            };
            return await dialog.ShowAsync() == ContentDialogResult.Primary;
        }
        finally
        {
            _dialogOpen = false;
        }
    }

    /// <summary>
    /// Puts <paramref name="query"/> in the query box and runs it (used by the Query History page). A request
    /// saved from the Advanced Query Rules page is opened there and run.
    /// </summary>
    public Task RunQueryAsync(string query)
    {
        if (AdvancedRequest.IsSaved(query))
        {
            if (!App.Settings.Data.ShowAdvancedPages) return Task.CompletedTask;
            NavView.SelectedItem = AdvancedSearchNavItem;
            if (ContentFrame.CurrentSourcePageType != typeof(AdvancedSearchPage)) ContentFrame.Navigate(typeof(AdvancedSearchPage));
            return ((AdvancedSearchPage)ContentFrame.Content).RunSavedAsync(query);
        }

        Search.QueryText = query;
        return RunSearchAsync();
    }

    private async Task RunSearchAsync()
    {
        QueryBox.IsSuggestionListOpen = false;
        // On the Advanced Query Rules page the query box is that page's query text, so searching runs its request.
        if (ContentFrame.Content is AdvancedSearchPage advanced)
        {
            await advanced.RunAsync();
            return;
        }
        // A search run from the Server Response page stays there, to show that search's details.
        if (ContentFrame.CurrentSourcePageType != typeof(HttpDetailsPage)) ShowSearchPage();
        await Search.SearchAsync();
    }

    public string SearchButtonGlyph(bool busy) => busy ? FluentIcons.Dismiss : FluentIcons.Search;

    public string SearchButtonLabel(bool busy) => busy ? "Cancel" : "Search";

    private void Search_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(SearchViewModel.Profile)) UpdateProfileDisplay();
    }

    #endregion

    #region Navigation

    private void NavView_ItemInvoked(NavigationView sender, NavigationViewItemInvokedEventArgs args)
    {
        switch ((args.InvokedItemContainer as NavigationViewItem)?.Tag as string)
        {

            case "search":
                ShowSearchPage();
                break;
            case "advanced":
                if (ContentFrame.CurrentSourcePageType != typeof(AdvancedSearchPage)) ContentFrame.Navigate(typeof(AdvancedSearchPage));
                break;
            case "http":
                if (ContentFrame.CurrentSourcePageType != typeof(HttpDetailsPage)) ContentFrame.Navigate(typeof(HttpDetailsPage));
                break;
            case "history":
                if (ContentFrame.CurrentSourcePageType != typeof(QueryHistoryPage)) ContentFrame.Navigate(typeof(QueryHistoryPage));
                break;
            case "profile":
                ShowProfileMenu();
                break;
            case "settings":
                ShowSettingsPage();
                break;
        }
    }

    /// <summary>Shows Primary Results (also used by the Advanced Query Rules page after it runs a query).</summary>
    public void ShowSearchPage()
    {
        NavView.SelectedItem = SearchNavItem;
        if (ContentFrame.CurrentSourcePageType != typeof(SearchPage)) ContentFrame.Navigate(typeof(SearchPage));
    }

    /// <summary>
    /// Shows or hides the Advanced Query Rules and Server Response pages (Settings › Search › Advanced pages),
    /// leaving for Search Results if one of them is open while they are hidden.
    /// </summary>
    public void ApplyAdvancedPagesVisibility()
    {
        bool show = App.Settings.Data.ShowAdvancedPages;
        AdvancedSearchNavItem.Visibility = HttpDetailsNavItem.Visibility = show ? Visibility.Visible : Visibility.Collapsed;
        if (!show && ContentFrame.Content is AdvancedSearchPage or HttpDetailsPage) ShowSearchPage();
    }

    private void ShowSettingsPage()
    {
        NavView.SelectedItem = SettingsNavItem;
        if (ContentFrame.CurrentSourcePageType != typeof(SettingsPage)) ContentFrame.Navigate(typeof(SettingsPage));
    }

    /// <summary>Opens Settings › Search › Favorite properties (the link in an empty Result Details summary).</summary>
    public void ShowFavoriteProperties()
    {
        ShowSettingsPage();
        (ContentFrame.Content as SettingsPage)?.OpenFavorites();
    }

    #endregion

    #region Profiles

    private void UpdateProfileDisplay()
    {
        var profile = App.Settings.CurrentProfile;
        WatchProfile(profile);

        ProfileNavItem.Content = profile?.EffectiveName ?? "Add a profile";
        ToolTipService.SetToolTip(ProfileNavItem, profile == null
            ? "Add a SharePoint site profile"
            : $"{profile.EffectiveName}\n{profile.SiteUrl}\n{profile.AccountDisplay}");
        Title = profile == null ? "SharePoint Search Explorer" : $"SharePoint Search Explorer – {profile.EffectiveName}";

        // Green check when signed in, yellow alert otherwise.
        bool signedIn = profile?.IsSignedIn == true;
        ApplyStatusBadge(ProfileBadge, signedIn);
        AutomationProperties.SetName(ProfileNavItem, signedIn ? "Profile, signed in" : "Profile, not signed in");
    }

    /// <summary>Follows the current profile so the badge updates the moment sign-in state changes.</summary>
    private void WatchProfile(ConnectionProfile? profile)
    {
        if (_watchedProfile == profile) return;
        if (_watchedProfile != null) _watchedProfile.PropertyChanged -= WatchedProfile_PropertyChanged;
        _watchedProfile = profile;
        if (profile != null) profile.PropertyChanged += WatchedProfile_PropertyChanged;
    }

    private void WatchedProfile_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName is not (nameof(ConnectionProfile.IsSignedIn) or nameof(ConnectionProfile.UserName) or nameof(ConnectionProfile.Name))) return;
        if (DispatcherQueue.HasThreadAccess) UpdateProfileDisplay();
        else DispatcherQueue.TryEnqueue(UpdateProfileDisplay);
    }

    /// <summary>
    /// Profile switcher: the current profile opens a submenu with its account actions; other profiles
    /// switch on click. "Add profile" comes last.
    /// </summary>
    private void ShowProfileMenu()
    {
        var settings = App.Settings;
        var current = settings.CurrentProfile;
        var menu = new MenuFlyout();

        foreach (var profile in settings.Data.Profiles)
        {
            if (profile == current)
            {
                menu.Items.Add(CurrentProfileMenuItem(profile));
                continue;
            }
            var item = MenuItem(profile.EffectiveName, FluentIcons.City, () =>
            {
                if (profile != settings.CurrentProfile) settings.SetCurrentProfile(profile);
            });
            ToolTipService.SetToolTip(item, $"Switch to this profile\n{profile.SiteUrl}");
            menu.Items.Add(item);
        }
        if (settings.Data.Profiles.Count > 0) menu.Items.Add(new MenuFlyoutSeparator());

        menu.Items.Add(MenuItem("Add profile…", FluentIcons.Add, async () => await EditProfileAsync(null)));

        menu.ShowAt(ProfileNavItem, new FlyoutShowOptions { Placement = FlyoutPlacementMode.RightEdgeAlignedBottom });
    }

    private MenuFlyoutSubItem CurrentProfileMenuItem(ConnectionProfile profile)
    {
        var sub = new MenuFlyoutSubItem { Text = profile.EffectiveName, Icon = FluentIcons.Create(FluentIcons.City) };
        ToolTipService.SetToolTip(sub, $"Current profile\n{profile.SiteUrl}");
        AutomationProperties.SetName(sub, $"{profile.EffectiveName}, current profile");
        AddCurrentProfileBadge(sub);

        sub.Items.Add(new MenuFlyoutItem
        {
            Text = profile.AccountDisplay,
            Icon = FluentIcons.Create(profile.IsSignedIn ? FluentIcons.CheckmarkCircle : FluentIcons.Warning),
            IsEnabled = false,
        });

        bool browser = profile.AuthMode == AuthMode.Browser;
        if (browser)
        {
            sub.Items.Add(profile.IsSignedIn
                ? MenuItem("Sign in with a different account…", FluentIcons.PersonSwap, async () => await Search.SignInAsync(chooseAccount: true))
                : MenuItem("Sign in", FluentIcons.PersonArrowRight, async () => await Search.SignInAsync(chooseAccount: false)));
        }
        sub.Items.Add(MenuItem("Edit profile…", FluentIcons.Edit, async () => await EditProfileAsync(profile)));
        if (browser && profile.IsSignedIn)
            sub.Items.Add(MenuItem("Sign out", FluentIcons.SignOut, async () => await Search.SignOutAsync()));

        sub.Items.Add(new MenuFlyoutSeparator());
        sub.Items.Add(MenuItem("Remove profile", FluentIcons.Delete, async () => await RemoveProfileAsync(profile)));
        return sub;
    }

    private static MenuFlyoutItem MenuItem(string text, string glyph, Action onClick)
    {
        var item = new MenuFlyoutItem { Text = text, Icon = FluentIcons.Create(glyph) };
        item.Click += (_, _) => onClick();
        return item;
    }

    /// <summary>
    /// Puts the nav rail's green check badge to the right of the profile name. MenuFlyoutSubItem only takes
    /// text, so the badge is added to its template, just before the submenu chevron.
    /// </summary>
    private static void AddCurrentProfileBadge(MenuFlyoutSubItem item)
    {
        item.Loaded += (_, _) =>
        {
            if (FindDescendant(item, "SubItemChevron") is not { Parent: Grid grid } chevron) return;
            if (grid.Children.OfType<InfoBadge>().Any()) return;

            var badge = CreateStatusBadge(signedIn: true);
            badge.Margin = new Thickness(12, 0, 0, 0);
            badge.HorizontalAlignment = HorizontalAlignment.Left;
            badge.VerticalAlignment = VerticalAlignment.Center;
            AutomationProperties.SetAccessibilityView(badge, AccessibilityView.Raw);
            Grid.SetColumn(badge, Grid.GetColumn(chevron));
            grid.Children.Add(badge);

            // Make room for the badge between the text and the chevron.
            var m = chevron.Margin;
            chevron.Margin = new Thickness(Math.Max(m.Left, 12 + badge.Width + 8), m.Top, m.Right, m.Bottom);
        };
    }

    private static FrameworkElement? FindDescendant(DependencyObject root, string name)
    {
        for (int i = 0; i < VisualTreeHelper.GetChildrenCount(root); i++)
        {
            var child = VisualTreeHelper.GetChild(root, i);
            if (child is FrameworkElement fe && fe.Name == name) return fe;
            if (FindDescendant(child, name) is { } found) return found;
        }
        return null;
    }

    /// <summary>The small badge on the profile icon: green check when signed in, yellow alert when not.</summary>
    private static InfoBadge CreateStatusBadge(bool signedIn)
    {
        var badge = new InfoBadge { Width = 12, Height = 12, MinWidth = 12, MinHeight = 12, Padding = new Thickness(0) };
        ApplyStatusBadge(badge, signedIn);
        return badge;
    }

    private static void ApplyStatusBadge(InfoBadge badge, bool signedIn)
    {
        badge.Style = (Style)Application.Current.Resources[signedIn ? "SuccessIconInfoBadgeStyle" : "CautionIconInfoBadgeStyle"];
        badge.IconSource = new FontIconSource
        {
            FontFamily = FluentIcons.FilledFontFamily,
            FontSize = 9,
            Glyph = signedIn ? FluentIcons.CheckmarkBadge : FluentIcons.ImportantBadge,
        };
    }

    /// <summary>Shows the add/edit dialog; a new or re-pointed profile becomes current and signs in.</summary>
    public async Task EditProfileAsync(ConnectionProfile? profile)
    {
        if (_dialogOpen) return;
        _dialogOpen = true;
        try
        {
            var dialog = new ProfileDialog(profile, App.Settings.Data.DefaultClientId) { XamlRoot = RootGrid.XamlRoot };
            dialog.RequestedTheme = RootGrid.ActualTheme;
            if (await dialog.ShowAsync() != ContentDialogResult.Primary || dialog.Result is not { } saved) return;

            bool connectionChanged = dialog.ConnectionChanged;
            if (connectionChanged && profile != null)
            {
                // Different site/tenant/app: the old account may not apply any more.
                saved.AccountId = "";
            }

            App.Settings.SaveProfile(saved);
            if (string.IsNullOrWhiteSpace(App.Settings.Data.DefaultClientId) && !string.IsNullOrWhiteSpace(saved.ClientId))
            {
                App.Settings.Data.DefaultClientId = saved.ClientId;
                App.Settings.Save();
            }

            if (App.Settings.CurrentProfile != saved || connectionChanged) App.Settings.SetCurrentProfile(saved);
            UpdateProfileDisplay();

            if (saved.AuthMode == AuthMode.Browser && (profile == null || connectionChanged))
                await Search.SignInAsync(chooseAccount: false);
        }
        finally
        {
            _dialogOpen = false;
        }
    }

    private async Task RemoveProfileAsync(ConnectionProfile profile)
    {
        var confirm = new ContentDialog
        {
            XamlRoot = RootGrid.XamlRoot,
            RequestedTheme = RootGrid.ActualTheme,
            Title = "Remove profile?",
            Content = $"“{profile.EffectiveName}” will be removed and its account signed out of this app.",
            PrimaryButtonText = "Remove",
            CloseButtonText = "Cancel",
            DefaultButton = ContentDialogButton.Close,
        };
        if (await confirm.ShowAsync() != ContentDialogResult.Primary) return;

        try
        {
            await App.Auth.SignOutAsync(profile);
        }
        catch (Exception)
        {
            // Removing the profile matters more than clearing its cached token.
        }
        App.Settings.RemoveProfile(profile);
        UpdateProfileDisplay();
    }

    /// <summary>
    /// Signs every profile out, forgets the embedded sign-in session and restores all settings to their
    /// defaults, then starts over like a first run.
    /// </summary>
    public async Task ResetToDefaultsAsync()
    {
        Search.CancelSearch();
        foreach (var profile in App.Settings.Data.Profiles.ToList())
        {
            try
            {
                await App.Auth.SignOutAsync(profile);
            }
            catch (Exception)
            {
                // The settings reset matters more than clearing a cached token.
            }
        }

        App.Settings.ResetToDefaults();
        WebViewSignIn.ClearBrowsingData();

        var data = App.Settings.Data;
        DisplayOptions.Current.WordWrap = data.WordWrap;
        DisplayOptions.Current.WrapCodeBlocks = data.WrapCodeBlocks;
        DisplayOptions.Current.UseFileTypeIcons = data.UseFileTypeIcons;
        DisplayOptions.Current.DenseResults = data.ResultDensity == "Dense";
        Search.ShowEmptyProperties = data.ShowEmptyProperties;
        Search.PropertiesHeader.SourceFilter = data.DefaultPropertySource;
        Search.RefreshSummary();
        ApplyTheme(data.Theme);
        ApplyAccentColor(data.AccentColor);
        Search.QueryText = "";
        UpdateProfileDisplay();
        ShowSearchPage();
        ApplyAdvancedPagesVisibility();

        await EditProfileAsync(null);
    }

    #endregion
}

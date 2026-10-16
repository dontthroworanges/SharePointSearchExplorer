using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Reflection;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using SharePointSearchExplorer.Models;
using SharePointSearchExplorer.Services;
using Windows.System;

namespace SharePointSearchExplorer.Views;

public sealed partial class SettingsPage : Page
{
    private bool _loading = true;
    private ConnectionProfile? _watchedProfile;

    public SettingsPage()
    {
        InitializeComponent();

        AdvancedPagesToggle.IsOn = Settings.ShowAdvancedPages;
        WrapCodeCard.IsEnabled = Settings.ShowAdvancedPages;
        RowLimitBox.Value = Settings.RowLimit;
        ThemeBox.SelectedIndex = Settings.Theme switch { "Light" => 1, "Dark" => 2, _ => 0 };
        LoadAccentColor();
        WordWrapToggle.IsOn = Settings.WordWrap;
        WrapCodeToggle.IsOn = Settings.WrapCodeBlocks;
        FileIconStyleBox.SelectedIndex = Settings.UseFileTypeIcons ? 1 : 0;
        DensityBox.SelectedIndex = Settings.ResultDensity == "Dense" ? 1 : 0;
        ShowEmptyToggle.IsOn = Settings.ShowEmptyProperties;
        PropertySourceBox.SelectedIndex = Settings.DefaultPropertySource switch
        {
            PropertyEntry.ManagedSource => 1,
            PropertyEntry.CrawledSource => 2,
            _ => 0,
        };
        AllowQueryDeletionToggle.IsOn = Settings.AllowQueryBoxDeletion;
        ConfirmQueryDeletionToggle.IsOn = Settings.ConfirmQueryBoxDeletion;
        ConfirmQueryDeletionToggle.IsEnabled = Settings.AllowQueryBoxDeletion;
        FillDelimiterBox(PairDelimiterBox, Settings.CopyPairDelimiter);
        FillDelimiterBox(AllDelimiterBox, Settings.CopyAllDelimiter);
        VersionCard.Header = $"SharePoint Search Explorer {Assembly.GetExecutingAssembly().GetName().Version?.ToString(3)}";
        DataFolderCard.Description = SettingsStore.DataFolder;
        UpdateProfile();

        CategoryNav.SelectedItem = CategoryNav.MenuItems[0];
        App.Settings.CurrentProfileChanged += OnCurrentProfileChanged;

        // Text fields bind TwoWay straight into AppSettings; persist when leaving the page.
        Unloaded += (_, _) =>
        {
            App.Settings.CurrentProfileChanged -= OnCurrentProfileChanged;
            WatchProfile(null);
            App.Settings.Save();
        };
        _loading = false;
    }

    public AppSettings Settings => App.Settings.Data;

    private void CategoryNav_SelectionChanged(NavigationView sender, NavigationViewSelectionChangedEventArgs args)
    {
        if (args.SelectedItem is NavigationViewItem item) ShowCategory(item);
    }

    /// <summary>Clicking the category that is already selected leaves its sub-page.</summary>
    private void CategoryNav_ItemInvoked(NavigationView sender, NavigationViewItemInvokedEventArgs args)
    {
        if (IsOnSubPage && args.InvokedItemContainer is NavigationViewItem item) ShowCategory(item);
    }

    private static Visibility Show(bool visible) => visible ? Visibility.Visible : Visibility.Collapsed;

    private bool IsOnSubPage => FavoritesPanel.Visibility == Visibility.Visible;

    /// <summary>Shows a category's settings, with its name as the page title.</summary>
    private void ShowCategory(NavigationViewItem item)
    {
        if (item.Tag is not string tag) return;
        CategoryTitle.Text = item.Content as string;
        CategoryTitle.Visibility = Visibility.Visible;
        SubPageBreadcrumb.Visibility = Visibility.Collapsed;
        FavoritesPanel.Visibility = Visibility.Collapsed;
        SearchPanel.Visibility = Show(tag == "Search");
        AccountsPanel.Visibility = Show(tag == "Accounts");
        AppearancePanel.Visibility = Show(tag == "Appearance");
        AboutPanel.Visibility = Show(tag == "About");
    }

    #region Favorite properties sub-page

    private const string FavoritesTitle = "Favorite properties";

    private readonly ObservableCollection<string> _availableProperties = new();
    private readonly ObservableCollection<string> _favoriteProperties = new();

    /// <summary>
    /// Opens the sub-page, like a Windows Settings row that leads to another page: the title becomes a
    /// breadcrumb ("Search > Favorite properties") whose first item goes back.
    /// </summary>
    private void FavoritesCard_Click(object sender, RoutedEventArgs e) => OpenFavorites();

    /// <summary>Shows the Favorite properties sub-page of the Search category (also opened from the Result Details summary).</summary>
    public void OpenFavorites()
    {
        CategoryNav.SelectedItem = CategoryNav.MenuItems[0];
        if (AvailableList.ItemsSource == null)
        {
            AvailableList.ItemsSource = _availableProperties;
            FavoritesList.ItemsSource = _favoriteProperties;
            // Also raised when a favorite is dragged to a new place.
            _favoriteProperties.CollectionChanged += (_, _) => UpdateFavoritesState();
        }
        AvailableFilterBox.Text = "";
        FavoritesStatusText.Text = "";
        LoadFavorites(Settings.FavoriteProperties);

        SubPageBreadcrumb.ItemsSource = new[] { CategoryTitle.Text, FavoritesTitle };
        SubPageBreadcrumb.Visibility = Visibility.Visible;
        CategoryTitle.Visibility = Visibility.Collapsed;
        SearchPanel.Visibility = Visibility.Collapsed;
        FavoritesPanel.Visibility = Visibility.Visible;
        ContentScroller.ChangeView(null, 0, null, disableAnimation: true);
    }

    private void SubPageBreadcrumb_ItemClicked(BreadcrumbBar sender, BreadcrumbBarItemClickedEventArgs args)
    {
        if (args.Index == 0 && CategoryNav.SelectedItem is NavigationViewItem item) ShowCategory(item);
    }

    private void LoadFavorites(IEnumerable<string> favorites)
    {
        _favoriteProperties.Clear();
        foreach (var name in favorites.Take(FavoritePropertyCatalog.MaxCount)) _favoriteProperties.Add(name);
        RefreshAvailableProperties();
    }

    /// <summary>The standard properties that aren't favorites and match the filter box.</summary>
    private void RefreshAvailableProperties()
    {
        var filter = AvailableFilterBox.Text.Trim();
        _availableProperties.Clear();
        foreach (var name in FavoritePropertyCatalog.Standard)
        {
            if (_favoriteProperties.Contains(name, StringComparer.OrdinalIgnoreCase)) continue;
            if (filter.Length > 0 && !name.Contains(filter, StringComparison.OrdinalIgnoreCase)) continue;
            _availableProperties.Add(name);
        }
        UpdateFavoritesState();
    }

    /// <summary>The slot counter and which buttons can be used.</summary>
    private void UpdateFavoritesState()
    {
        int free = FavoritePropertyCatalog.MaxCount - _favoriteProperties.Count;
        FavoriteSlotsText.Text = free == 1
            ? $"1 of {FavoritePropertyCatalog.MaxCount} slots available"
            : $"{Math.Max(free, 0)} of {FavoritePropertyCatalog.MaxCount} slots available";
        AddFavoriteButton.IsEnabled = free > 0 && AvailableList.SelectedItems.Count > 0;
        RemoveFavoriteButton.IsEnabled = FavoritesList.SelectedItems.Count > 0;
        ClearFavoritesButton.IsEnabled = _favoriteProperties.Count > 0;
        bool changed = !_favoriteProperties.SequenceEqual(Settings.FavoriteProperties);
        SaveFavoritesButton.IsEnabled = changed;
        if (changed) FavoritesStatusText.Text = "";
    }

    private void AvailableFilterBox_TextChanged(object sender, TextChangedEventArgs e) => RefreshAvailableProperties();

    /// <summary>
    /// Only one list has a selection at a time, so only one of the add and remove arrows is active: selecting in
    /// one list clears the other. Clearing raises the other list's event with only removed items, so it doesn't loop.
    /// </summary>
    private void FavoriteLists_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (e.AddedItems.Count > 0)
        {
            var other = sender == AvailableList ? FavoritesList : AvailableList;
            if (other.SelectedItems.Count > 0) other.SelectedItems.Clear();
        }
        UpdateFavoritesState();
    }

    private void AddFavorite_Click(object sender, RoutedEventArgs e) => AddFavorites(AvailableList.SelectedItems.OfType<string>().ToList());

    private void AvailableList_DoubleTapped(object sender, Microsoft.UI.Xaml.Input.DoubleTappedRoutedEventArgs e)
    {
        if (DoubleTappedProperty(e) is { } name) AddFavorites([name]);
    }

    /// <summary>The property of the list item that was double-clicked; null for the list's empty space.</summary>
    private static string? DoubleTappedProperty(Microsoft.UI.Xaml.Input.DoubleTappedRoutedEventArgs e)
    {
        for (var element = e.OriginalSource as DependencyObject; element != null; element = Microsoft.UI.Xaml.Media.VisualTreeHelper.GetParent(element))
            if (element is ListViewItem item) return item.Content as string;
        return null;
    }

    /// <summary>Moves properties to the favorites, as far as there are free slots.</summary>
    private void AddFavorites(IReadOnlyList<string> names)
    {
        foreach (var name in names)
        {
            if (_favoriteProperties.Count >= FavoritePropertyCatalog.MaxCount) break;
            _availableProperties.Remove(name);
            _favoriteProperties.Add(name);
        }
        UpdateFavoritesState();
    }

    private void RemoveFavorite_Click(object sender, RoutedEventArgs e) => RemoveFavorites(FavoritesList.SelectedItems.OfType<string>().ToList());

    private void FavoritesList_DoubleTapped(object sender, Microsoft.UI.Xaml.Input.DoubleTappedRoutedEventArgs e)
    {
        if (DoubleTappedProperty(e) is { } name) RemoveFavorites([name]);
    }

    private void RemoveFavorites(IReadOnlyList<string> names)
    {
        foreach (var name in names) _favoriteProperties.Remove(name);
        // Puts them back in the available list at their sorted place.
        RefreshAvailableProperties();
    }

    /// <summary>Empties the favorites; like any other change it is kept by "Save Favorites".</summary>
    private void ClearFavorites_Click(object sender, RoutedEventArgs e) => LoadFavorites([]);

    /// <summary>Fills the favorites with the defaults; like any other change it is kept by "Save Favorites".</summary>
    private void ResetFavorites_Click(object sender, RoutedEventArgs e) => LoadFavorites(FavoritePropertyCatalog.Defaults);

    private void SaveFavorites_Click(object sender, RoutedEventArgs e)
    {
        Settings.FavoriteProperties = _favoriteProperties.ToList();
        App.Settings.Save();
        // The summary of the selected result changes straight away; new searches also request the favorites.
        App.Search.RefreshSummary();
        UpdateFavoritesState();
        FavoritesStatusText.Text = "Favorites saved";
    }

    #endregion

    /// <summary>Settings content: at most 1000px of cards (plus padding), never wider than the visible area.</summary>
    private const double MaxContentWidth = 1000 + 72;

    private void ContentScroller_SizeChanged(object sender, SizeChangedEventArgs e) =>
        CenterContent(ContentPanel, e.NewSize.Width);

    /// <summary>
    /// Like Windows 11 Settings: the column of cards is capped at <see cref="MaxContentWidth"/> and sits in the
    /// middle of the page when there is more room than that. Both the width and the left margin are set from
    /// the viewport's width, so nothing depends on the panel being re-measured (see the note in the XAML).
    /// Also used by the Query History page.
    /// </summary>
    internal static void CenterContent(FrameworkElement panel, double viewportWidth)
    {
        panel.Width = Math.Min(viewportWidth, MaxContentWidth);
        panel.Margin = new Thickness(Math.Max(0, Math.Floor((viewportWidth - MaxContentWidth) / 2)), 0, 0, 0);
    }

    private void AdvancedPagesToggle_Toggled(object sender, RoutedEventArgs e)
    {
        if (_loading) return;
        Settings.ShowAdvancedPages = AdvancedPagesToggle.IsOn;
        // The Server Response code blocks setting only matters while that page can be opened.
        WrapCodeCard.IsEnabled = AdvancedPagesToggle.IsOn;
        App.MainWindow.ApplyAdvancedPagesVisibility();
        App.Settings.Save();
    }

    private void RowLimitBox_ValueChanged(NumberBox sender, NumberBoxValueChangedEventArgs args)
    {
        if (_loading || double.IsNaN(args.NewValue)) return;
        Settings.RowLimit = (int)Math.Clamp(args.NewValue, 1, 500);
    }

    private void ThemeBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_loading || ThemeBox.SelectedItem is not ComboBoxItem { Tag: string theme }) return;
        Settings.Theme = theme;
        App.MainWindow.ApplyTheme(theme);
        App.Settings.Save();
    }

    private void LoadAccentColor()
    {
        bool custom = Helpers.AccentColors.Parse(Settings.AccentColor) != null;
        AccentModeBox.SelectedIndex = custom ? 1 : 0;
        AccentPalettePanel.Visibility = custom ? Visibility.Visible : Visibility.Collapsed;
        AccentGrid.SelectedItem = Helpers.AccentColors.Palette.FirstOrDefault(s =>
            string.Equals(s.Hex, Settings.AccentColor, StringComparison.OrdinalIgnoreCase));
    }

    private void AccentModeBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_loading) return;
        bool custom = AccentModeBox.SelectedIndex == 1;
        AccentPalettePanel.Visibility = custom ? Visibility.Visible : Visibility.Collapsed;
        if (!custom)
        {
            SetAccentColor("");
        }
        else
        {
            // Start from the Windows accent colour when it is one of the swatches, like Windows Settings does.
            var system = new Windows.UI.ViewManagement.UISettings().GetColorValue(Windows.UI.ViewManagement.UIColorType.Accent);
            var swatch = AccentGrid.SelectedItem as Helpers.AccentSwatch
                ?? Helpers.AccentColors.Palette.FirstOrDefault(s => s.Color == system)
                ?? Helpers.AccentColors.Palette.First(s => s.Hex == "#0078D4");
            AccentGrid.SelectedItem = swatch;
            SetAccentColor(swatch.Hex);
        }
    }

    private void AccentGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_loading || AccentGrid.SelectedItem is not Helpers.AccentSwatch swatch) return;
        SetAccentColor(swatch.Hex);
    }

    private void SetAccentColor(string hex)
    {
        if (string.Equals(Settings.AccentColor, hex, StringComparison.OrdinalIgnoreCase)) return;
        Settings.AccentColor = hex;
        App.MainWindow.ApplyAccentColor(hex);
        App.Settings.Save();
    }

    private static void FillDelimiterBox(ComboBox box, string selectedKey)
    {
        foreach (var option in Helpers.Delimiters.All)
        {
            var item = new ComboBoxItem { Content = option.Label, Tag = option.Key };
            box.Items.Add(item);
            if (string.Equals(option.Key, selectedKey, StringComparison.OrdinalIgnoreCase)) box.SelectedItem = item;
        }
        if (box.SelectedIndex < 0) box.SelectedIndex = 0;
    }

    private void DelimiterBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_loading || sender is not ComboBox { SelectedItem: ComboBoxItem { Tag: string key } } box) return;
        if (box == PairDelimiterBox) Settings.CopyPairDelimiter = key;
        else Settings.CopyAllDelimiter = key;
        App.Settings.Save();
    }

    private void FileIconStyleBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_loading) return;
        Settings.UseFileTypeIcons = FileIconStyleBox.SelectedIndex == 1;
        Helpers.DisplayOptions.Current.UseFileTypeIcons = Settings.UseFileTypeIcons;
        App.Settings.Save();
    }

    private void WordWrapToggle_Toggled(object sender, RoutedEventArgs e)
    {
        if (_loading) return;
        Settings.WordWrap = WordWrapToggle.IsOn;
        Helpers.DisplayOptions.Current.WordWrap = WordWrapToggle.IsOn;
        App.Settings.Save();
    }

    private void WrapCodeToggle_Toggled(object sender, RoutedEventArgs e)
    {
        if (_loading) return;
        Settings.WrapCodeBlocks = WrapCodeToggle.IsOn;
        Helpers.DisplayOptions.Current.WrapCodeBlocks = WrapCodeToggle.IsOn;
        App.Settings.Save();
    }

    private void AllowQueryDeletionToggle_Toggled(object sender, RoutedEventArgs e)
    {
        // Confirmation only matters while deleting is allowed.
        ConfirmQueryDeletionToggle.IsEnabled = AllowQueryDeletionToggle.IsOn;
        if (_loading) return;
        Settings.AllowQueryBoxDeletion = AllowQueryDeletionToggle.IsOn;
        App.Settings.Save();
    }

    private void ConfirmQueryDeletionToggle_Toggled(object sender, RoutedEventArgs e)
    {
        if (_loading) return;
        Settings.ConfirmQueryBoxDeletion = ConfirmQueryDeletionToggle.IsOn;
        App.Settings.Save();
    }

    private void ShowEmptyToggle_Toggled(object sender, RoutedEventArgs e)
    {
        if (_loading) return;
        Settings.ShowEmptyProperties = ShowEmptyToggle.IsOn;
        // A new default applies straight away, replacing any change made with a table's toggle.
        App.Search.ShowEmptyProperties = ShowEmptyToggle.IsOn;
        App.Settings.Save();
    }

    private void PropertySourceBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_loading || PropertySourceBox.SelectedItem is not ComboBoxItem item) return;
        var source = item.Tag as string ?? "";
        Settings.DefaultPropertySource = source;
        // A new default applies straight away, replacing any change made with the table's Source header.
        App.Search.PropertiesHeader.SourceFilter = source;
        App.Settings.Save();
    }

    private void DensityBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_loading || DensityBox.SelectedItem is not ComboBoxItem { Tag: string density }) return;
        Settings.ResultDensity = density;
        Helpers.DisplayOptions.Current.DenseResults = density == "Dense";
        App.Settings.Save();
    }

    private async void ResetApp_Click(object sender, RoutedEventArgs e)
    {
        var confirm = new ContentDialog
        {
            XamlRoot = XamlRoot,
            RequestedTheme = ActualTheme,
            Title = "Reset app to defaults?",
            Content = "All profiles will be removed and signed out, query history will be cleared and every setting will return to its default. This can't be undone.",
            PrimaryButtonText = "Reset",
            CloseButtonText = "Cancel",
            DefaultButton = ContentDialogButton.Close,
        };
        if (await confirm.ShowAsync() != ContentDialogResult.Primary) return;
        await App.MainWindow.ResetToDefaultsAsync();
    }

    #region Profile

    private async void EditProfile_Click(object sender, RoutedEventArgs e) =>
        await App.MainWindow.EditProfileAsync(App.Settings.CurrentProfile);

    private async void AddProfile_Click(object sender, RoutedEventArgs e) =>
        await App.MainWindow.EditProfileAsync(null);

    private async void OpenDataFolder_Click(object sender, RoutedEventArgs e)
    {
        Directory.CreateDirectory(SettingsStore.DataFolder);
        await Launcher.LaunchFolderPathAsync(SettingsStore.DataFolder);
    }

    private void OnCurrentProfileChanged(object? sender, EventArgs e) => UpdateProfile();

    private void UpdateProfile()
    {
        var profile = App.Settings.CurrentProfile;
        WatchProfile(profile);

        ProfilePicture.DisplayName = profile?.EffectiveName ?? "";
        ProfileNameText.Text = profile?.EffectiveName ?? "No profile";
        ProfileAccountText.Text = profile?.AccountDisplay ?? "Add a SharePoint site to get started";

        CurrentProfileCard.Header = profile?.EffectiveName ?? "No profile yet";
        CurrentProfileCard.Description = profile == null ? "Add a SharePoint site to start searching." : $"{profile.SiteUrl}\n{profile.AccountDisplay}";
        EditProfileButton.IsEnabled = profile != null;
    }

    private void WatchProfile(ConnectionProfile? profile)
    {
        if (_watchedProfile == profile) return;
        if (_watchedProfile != null) _watchedProfile.PropertyChanged -= WatchedProfile_PropertyChanged;
        _watchedProfile = profile;
        if (profile != null) profile.PropertyChanged += WatchedProfile_PropertyChanged;
    }

    private void WatchedProfile_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (DispatcherQueue.HasThreadAccess) UpdateProfile();
        else DispatcherQueue.TryEnqueue(UpdateProfile);
    }

    #endregion
}

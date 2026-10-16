using System.ComponentModel;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Controls.Tabular;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Media.Imaging;
using SharePointSearchExplorer.Helpers;
using SharePointSearchExplorer.Models;
using SharePointSearchExplorer.ViewModels;
using Windows.ApplicationModel.DataTransfer;
using Windows.System;

namespace SharePointSearchExplorer.Views;

public sealed partial class SearchPage : Page
{
    public SearchPage()
    {
        InitializeComponent();
        PropertyNameColumn.Header = ViewModel.PropertiesHeader;
        PropertySourceColumn.Header = ViewModel.PropertiesHeader;
        PropertyStackedColumn.Header = ViewModel.PropertiesHeader;
        UpdateResultsSource();
        ViewModel.PropertyChanged += ViewModel_PropertyChanged;
    }

    public SearchViewModel ViewModel => App.Search;

    public Visibility IsEmpty(int count) => count == 0 ? Visibility.Visible : Visibility.Collapsed;

    public bool IsNot(bool value) => !value;

    // Result icon style. Items without a file type icon (e.g. people) keep the outline glyph in both styles.
    public static Visibility ShowGlyph(string? fileTypeIcon, bool useFileTypeIcons) =>
        useFileTypeIcons && fileTypeIcon != null ? Visibility.Collapsed : Visibility.Visible;

    public static Visibility ShowFileTypeImage(string? fileTypeIcon, bool useFileTypeIcons) =>
        useFileTypeIcons && fileTypeIcon != null ? Visibility.Visible : Visibility.Collapsed;

    // A favorite property in the summary: a dash when the result has no value for it (and then no tooltip).
    public static string FactText(string value) => value.Length == 0 ? "—" : value;

    public static object? FactTip(string value) => value.Length == 0 ? null : value;

    // Show/hide empty properties button (also used by PropertiesWindow): labelled with what a click does.
    public static string EmptyToggleLabel(bool showEmpty) => showEmpty ? "Hide empty" : "Show empty";

    public static string EmptyToggleGlyph(bool showEmpty) => showEmpty ? FluentIcons.EyeOff : FluentIcons.Eye;

    public static string EmptyToggleTip(bool showEmpty) => showEmpty ? "Hide properties with no value" : "Show properties with no value";

    private static readonly Dictionary<string, SvgImageSource> s_fileTypeImages = new();

    /// <summary>Unmodified Fluent file type SVG, shared between rows; not loaded until colour icons are on.</summary>
    public static ImageSource? FileTypeImage(string? fileTypeIcon, bool useFileTypeIcons)
    {
        if (!useFileTypeIcons || fileTypeIcon == null) return null;
        if (!s_fileTypeImages.TryGetValue(fileTypeIcon, out var image))
        {
            image = new SvgImageSource(new Uri($"ms-appx:///Assets/FileTypes/{fileTypeIcon}.svg"));
            s_fileTypeImages[fileTypeIcon] = image;
        }
        return image;
    }


    private void ViewModel_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(SearchViewModel.SelectedResult) && ViewModel.SelectedResult != null)
            ResultsList.ScrollIntoView(ViewModel.SelectedResult);
        else if (e.PropertyName == nameof(SearchViewModel.ResultGroups))
            UpdateResultsSource();
    }

    /// <summary>Shows the flat results, or the groups while grouping is on, and keeps the selection.</summary>
    private void UpdateResultsSource()
    {
        if (ViewModel.IsGrouped)
        {
            GroupedResults.Source = ViewModel.ResultGroups;
            ResultsList.ItemsSource = GroupedResults.View;
        }
        else if (ResultsList.ItemsSource != ViewModel.Results)
        {
            ResultsList.ItemsSource = ViewModel.Results;
        }
        // A new source clears the list's selection; the detail pane still shows the selected result.
        if (ViewModel.SelectedResult is { } selected)
        {
            ResultsList.SelectedItem = selected;
            ResultsList.ScrollIntoView(selected);
        }
    }

    private void ResultsList_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (ResultsList.SelectedItem is SearchResultItem item) ViewModel.SelectedResult = item;
    }

    /// <summary>The result whose right-click menu is open; the menu's actions apply to it, not to the selection.</summary>
    private SearchResultItem? _menuItem;

    /// <summary>
    /// Right-click (or Shift+F10 / menu key) on a result: shows its actions, disabling those it has no URL
    /// for (e.g. no ParentLink). The selection, and so the details pane, stays as it is.
    /// </summary>
    private void ResultsList_ContextRequested(UIElement sender, ContextRequestedEventArgs args)
    {
        // The row is the list item above the element under the pointer; group headers have none.
        SearchResultItem? item = null;
        for (var element = args.OriginalSource as DependencyObject; element != null && item == null; element = VisualTreeHelper.GetParent(element))
            item = (element as ListViewItem)?.Content as SearchResultItem;
        if (item == null) return;
        args.Handled = true;

        PrepareResultMenu(item);
        if (args.TryGetPosition(sender, out var point))
            ResultMenu.ShowAt(sender, new FlyoutShowOptions { Position = point });
        else
            ResultMenu.ShowAt((FrameworkElement)(ResultsList.ContainerFromItem(item) ?? ResultsList));
    }

    /// <summary>The Result options button beside the title in Result Details: the selected result's actions.</summary>
    private void ResultOptions_Click(object sender, RoutedEventArgs e)
    {
        if (ViewModel.SelectedResult is not { } item) return;
        PrepareResultMenu(item);
        ResultMenu.ShowAt((FrameworkElement)sender, new FlyoutShowOptions { Placement = FlyoutPlacementMode.BottomEdgeAlignedRight });
    }

    /// <summary>Points the result menu at <paramref name="item"/>, disabling actions it has no URL for.</summary>
    private void PrepareResultMenu(SearchResultItem item)
    {
        _menuItem = item;
        OpenItem.IsEnabled = item.OpenUri != null;
        OpenParentItem.IsEnabled = item.ParentUri != null;
    }

    private void ErrorBar_Closed(InfoBar sender, InfoBarClosedEventArgs args) => ViewModel.ErrorMessage = null;

    #region Toolbar

    private void SortFlyout_Opening(object sender, object e)
    {
        foreach (var item in SortFlyout.Items.OfType<RadioMenuFlyoutItem>())
            item.IsChecked = item.Tag as string == ViewModel.Sort.ToString();
    }

    private async void Sort_Click(object sender, RoutedEventArgs e)
    {
        if (sender is RadioMenuFlyoutItem { Tag: string tag } && Enum.TryParse<SortOption>(tag, out var sort))
            await ViewModel.SetSortAsync(sort);
    }

    private void GroupFlyout_Opening(object sender, object e)
    {
        foreach (var item in GroupFlyout.Items.OfType<RadioMenuFlyoutItem>())
            item.IsChecked = item.Tag as string == ViewModel.GroupBy.ToString();
    }

    private void Group_Click(object sender, RoutedEventArgs e)
    {
        if (sender is RadioMenuFlyoutItem { Tag: string tag } && Enum.TryParse<GroupOption>(tag, out var group))
            ViewModel.GroupBy = group;
    }

    private async void ApplyFilter_Click(object sender, RoutedEventArgs e)
    {
        FilterFlyout.Hide();
        await ViewModel.ApplyFiltersAsync();
    }

    private async void ClearFilter_Click(object sender, RoutedEventArgs e)
    {
        FilterFlyout.Hide();
        await ViewModel.ClearFiltersAsync();
    }

    private async void LoadMore_Click(object sender, RoutedEventArgs e) => await ViewModel.LoadMoreAsync();

    #endregion

    #region Details actions

    // Open Result, Open Parent Location and Copy link: the result menu (right-click, which need not be the selected
    // result, or the Result options button).
    private async void Open_Click(object sender, RoutedEventArgs e)
    {
        if (_menuItem?.OpenUri is { } uri) await Launcher.LaunchUriAsync(uri);
    }

    private async void OpenParent_Click(object sender, RoutedEventArgs e)
    {
        if (_menuItem?.ParentUri is { } uri) await Launcher.LaunchUriAsync(uri);
    }

    private void CopyLink_Click(object sender, RoutedEventArgs e)
    {
        if (_menuItem is { Path.Length: > 0 } item) CopyText(item.Path);
    }

    private void AddFavorites_Click(object sender, RoutedEventArgs e) => App.MainWindow.ShowFavoriteProperties();

    private void ToggleEmpty_Click(object sender, RoutedEventArgs e) => ViewModel.ShowEmptyProperties = !ViewModel.ShowEmptyProperties;

    private void CopyProperties_Click(object sender, RoutedEventArgs e) => CopyText(ViewModel.GetVisiblePropertiesText());

    private async void ReloadProperties_Click(object sender, RoutedEventArgs e) => await ViewModel.ReloadPropertiesAsync();

    /// <summary>
    /// A click on the Property header (the only sortable one) reverses the sort. The view model sorts the rows,
    /// so the TableView's own sort is cancelled.
    /// </summary>
    private void PropertiesGrid_Sorting(TableView sender, TableViewSortingEventArgs e)
    {
        e.Cancel = true;
        ViewModel.PropertiesHeader.SortDescending = !ViewModel.PropertiesHeader.SortDescending;
    }

    // Null until the first layout, which also removes the stacked column the wide layout doesn't use.
    private bool? _stacked;

    /// <summary>
    /// Narrow pane: the table stacks each name above its value, and the pane scrolls as a whole. The table is
    /// capped at the pane's height, so it keeps its own scrolling (and row virtualization) inside it.
    /// </summary>
    private void DetailScroller_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        double tableWidth = e.NewSize.Width - DetailScroller.Padding.Right;
        bool stacked = tableWidth < PropertyGridLayout.StackBelowWidth;
        PropertiesGrid.MaxHeight = stacked ? e.NewSize.Height : double.PositiveInfinity;
        if (stacked == _stacked) return;

        _stacked = stacked;
        PropertyGridLayout.SetStacked(PropertiesGrid, stacked);
        PropertyGridLayout.SetFilterBarStacked(PropertiesFilterBar, stacked);
        DetailScroller.VerticalScrollBarVisibility = stacked ? ScrollBarVisibility.Auto : ScrollBarVisibility.Disabled;
        if (!stacked) DetailScroller.ChangeView(null, 0, null, disableAnimation: true);
    }

    private void PropertyRow_ContextRequested(UIElement sender, ContextRequestedEventArgs args) =>
        PropertyClipboard.ShowRowMenu(sender, args);

    private static void CopyText(string text) => PropertyClipboard.CopyText(text);

    /// <summary>
    /// Opens the property table in its own window, with the same filters and sort, as a snapshot that later
    /// searches don't change.
    /// </summary>
    private void PopOut_Click(object sender, RoutedEventArgs e)
    {
        if (ViewModel.SelectedResult is not { } item || ViewModel.GetSelectedProperties() is not { } rows) return;
        PropertiesWindow.Show(item, rows, ViewModel.PropertyFilter, ViewModel.ShowEmptyProperties, ViewModel.PropertiesHeader);
    }

    #endregion
}

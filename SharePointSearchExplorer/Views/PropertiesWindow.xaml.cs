using System.Collections.ObjectModel;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Tabular;
using Microsoft.UI.Xaml.Input;
using SharePointSearchExplorer.Models;
using Windows.Graphics;

namespace SharePointSearchExplorer.Views;

/// <summary>
/// The Result Details property table in its own window. It keeps a snapshot of the rows, so new searches
/// don't change it and several items can be compared side by side.
/// </summary>
public sealed partial class PropertiesWindow : Window
{
    private const string Subtitle = "Properties";

    private static readonly List<PropertiesWindow> s_open = new();

    private readonly IReadOnlyList<PropertyEntry> _rows;
    private readonly ObservableCollection<PropertyEntry> _visible = new();
    private bool _showEmpty;

    /// <summary>The table's header state (count, sort and Source filter), as on the Search Results page.</summary>
    private readonly ColumnHeaderInfo _header;

    private PropertiesWindow(SearchResultItem item, IReadOnlyList<PropertyEntry> rows, string filter, bool showEmpty, ColumnHeaderInfo header)
    {
        InitializeComponent();
        _rows = rows.ToList();
        // A copy: the window starts with the page's sort and Source filter, then keeps its own.
        _header = new ColumnHeaderInfo(header.Title) { SortDescending = header.SortDescending, SourceFilter = header.SourceFilter };

        Title = $"{item.Title} – {Subtitle}";
        ExtendsContentIntoTitleBar = true;
        SetTitleBar(WindowTitleBar);
        WindowTitleBar.Title = item.Title;
        WindowTitleBar.Subtitle = Subtitle;
        AppWindow.SetIcon(Path.Combine(AppContext.BaseDirectory, "Assets", "AppIcon.ico"));
        PathText.Text = item.Path;
        PathText.Visibility = item.Path.Length > 0 ? Visibility.Visible : Visibility.Collapsed;

        foreach (var column in Columns()) PropertyGrid.Columns.Add(column);
        PropertyGrid.ItemsSource = _visible;

        _showEmpty = showEmpty;
        UpdateEmptyToggle();
        FilterBox.Text = filter;
        Refresh();
        _header.PropertyChanged += (_, e) =>
        {
            if (e.PropertyName is nameof(ColumnHeaderInfo.SortDescending) or nameof(ColumnHeaderInfo.SourceFilter)) Refresh();
        };

        ApplyTheme(App.Settings.Data.Theme);
        Closed += (_, _) => s_open.Remove(this);
    }

    /// <summary>Opens a new window for the given rows, cascaded over the main window.</summary>
    public static void Show(SearchResultItem item, IReadOnlyList<PropertyEntry> rows, string filter, bool showEmpty, ColumnHeaderInfo header)
    {
        var window = new PropertiesWindow(item, rows, filter, showEmpty, header);
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

    /// <summary>
    /// Same columns as the table on the Search Results page; every header shows <see cref="_header"/>. The last
    /// column is the stacked one that replaces the others in a narrow window (see PropertyGridLayout).
    /// </summary>
    private IEnumerable<TableViewColumn> Columns()
    {
        var stacked = Column(_header, null, "StackedCellTemplate", "StackedColumnHeaderTemplate", canSort: true);
        stacked.CanResize = false;
        stacked.Visibility = Visibility.Collapsed;
        return
        [
            Column(_header, 320, "NameCellTemplate", "CountColumnHeaderTemplate", canSort: true),
            Column(_header, 110, "SourceCellTemplate", "SourceColumnHeaderTemplate"),
            Column("Value", null, "ValueCellTemplate", "ValueColumnHeaderTemplate"),
            stacked,
        ];

        TableViewTemplateColumn Column(object header, double? width, string template, string headerTemplate, bool canSort = false)
        {
            var column = new TableViewTemplateColumn
            {
                Header = header,
                HeaderTemplate = (DataTemplate)RootGrid.Resources[headerTemplate],
                Width = width is { } w ? new GridLength(w) : new GridLength(1, GridUnitType.Star),
                CellTemplate = (DataTemplate)RootGrid.Resources[template],
                CanSort = canSort,
            };
            if (canSort) column.SortMemberPath = nameof(PropertyEntry.Name);
            return column;
        }
    }

    /// <summary>
    /// A click on the Property header (the only sortable one) reverses the sort. The window sorts the rows
    /// itself, so the TableView's own sort is cancelled.
    /// </summary>
    private void PropertyGrid_Sorting(TableView sender, TableViewSortingEventArgs e)
    {
        e.Cancel = true;
        _header.SortDescending = !_header.SortDescending;
    }

    // Null until the first layout, which also removes the stacked column the wide layout doesn't use.
    private bool? _stacked;

    private void PropertyGrid_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        bool stacked = e.NewSize.Width < PropertyGridLayout.StackBelowWidth;
        if (stacked == _stacked) return;
        _stacked = stacked;
        PropertyGridLayout.SetStacked(PropertyGrid, stacked);
        PropertyGridLayout.SetFilterBarStacked(FilterBar, stacked);
    }

    private void FilterBox_TextChanged(object sender, TextChangedEventArgs e) => Refresh();

    private void EmptyToggle_Click(object sender, RoutedEventArgs e)
    {
        _showEmpty = !_showEmpty;
        UpdateEmptyToggle();
        Refresh();
    }

    private void UpdateEmptyToggle()
    {
        EmptyToggle.Label = SearchPage.EmptyToggleLabel(_showEmpty);
        EmptyToggleIcon.Glyph = SearchPage.EmptyToggleGlyph(_showEmpty);
        ToolTipService.SetToolTip(EmptyToggle, SearchPage.EmptyToggleTip(_showEmpty));
    }

    private void Refresh() => _header.Fill(_visible, _rows, FilterBox.Text, _showEmpty);

    private void CopyAll_Click(object sender, RoutedEventArgs e) =>
        PropertyClipboard.CopyText(PropertyClipboard.JoinRows(_visible));

    private void PropertyGrid_ContextRequested(UIElement sender, ContextRequestedEventArgs args) =>
        PropertyClipboard.ShowRowMenu(sender, args);
}

using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Tabular;

namespace SharePointSearchExplorer.Views;

/// <summary>
/// Narrow layout for property tables, shared by the Search Results page and pop-out property windows: below
/// <see cref="StackBelowWidth"/> the table shows one column with each name above its value.
/// </summary>
internal static class PropertyGridLayout
{
    /// <summary>Table width below which the side-by-side columns leave too little room for values.</summary>
    public const double StackBelowWidth = 600;

    // Every column a table was declared with, captured the first time its layout is set.
    private static readonly System.Runtime.CompilerServices.ConditionalWeakTable<TableView, List<TableViewColumn>> s_columns = new();

    /// <summary>
    /// Switches between the side-by-side columns and the stacked column, which is the table's last column.
    /// The unused columns are removed, not collapsed, so rows don't build cells for columns they don't show.
    /// The column objects are kept, so widths the user set by dragging come back with them.
    /// </summary>
    public static void SetStacked(TableView grid, bool stacked)
    {
        var all = s_columns.GetValue(grid, g => g.Columns.ToList());
        var stackedColumn = all[^1];
        var wanted = all.Where(c => (c == stackedColumn) == stacked).ToList();
        if (grid.Columns.SequenceEqual(wanted)) return;

        grid.Columns.Clear();
        foreach (var column in wanted)
        {
            column.Visibility = Visibility.Visible;
            grid.Columns.Add(column);
        }
    }

    /// <summary>
    /// The filter bar above a table (filter box, spacer, buttons). Stacked, the filter box moves to a row of
    /// its own under the buttons and takes the full width; buttons that still don't fit go to the "more" menu.
    /// </summary>
    public static void SetFilterBarStacked(Grid bar, bool stacked)
    {
        if (bar.RowDefinitions.Count == 0)
        {
            bar.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            bar.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        }
        bar.RowSpacing = stacked ? 4 : 0;
        foreach (var child in bar.Children.OfType<FrameworkElement>())
        {
            bool isBox = child is TextBox;
            Grid.SetRow(child, stacked && isBox ? 1 : 0);
            Grid.SetColumn(child, stacked || isBox ? 0 : 2);
            Grid.SetColumnSpan(child, stacked ? 3 : 1);
            if (child is CommandBar commands)
                commands.OverflowButtonVisibility = stacked ? CommandBarOverflowButtonVisibility.Auto : CommandBarOverflowButtonVisibility.Collapsed;
        }
    }
}

using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using SharePointSearchExplorer.Models;

namespace SharePointSearchExplorer.Views;

/// <summary>Code-behind so the cell and header templates can use x:Bind.</summary>
public sealed partial class PropertyGridResources : ResourceDictionary
{
    public PropertyGridResources() => InitializeComponent();

    /// <summary>The Source header's menu: show every row, or only the managed or the crawled properties.</summary>
    private void SourceFilter_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { DataContext: ColumnHeaderInfo header } button) return;

        var menu = new MenuFlyout { Placement = FlyoutPlacementMode.BottomEdgeAlignedLeft };
        menu.Items.Add(Item("All", ""));
        menu.Items.Add(new MenuFlyoutSeparator());
        menu.Items.Add(Item("Managed property", PropertyEntry.ManagedSource));
        menu.Items.Add(Item("Crawled property", PropertyEntry.CrawledSource));
        menu.ShowAt(button);

        RadioMenuFlyoutItem Item(string text, string source)
        {
            var item = new RadioMenuFlyoutItem { Text = text, GroupName = "source", IsChecked = header.SourceFilter == source };
            item.Click += (_, _) => header.SourceFilter = source;
            return item;
        }
    }
}

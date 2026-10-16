using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Input;
using SharePointSearchExplorer.Helpers;
using SharePointSearchExplorer.Models;
using Windows.ApplicationModel.DataTransfer;

namespace SharePointSearchExplorer.Views;

/// <summary>Clipboard actions for property tables, shared by the Search Results page and pop-out property windows.</summary>
internal static class PropertyClipboard
{
    public static void CopyText(string text)
    {
        var package = new DataPackage();
        package.SetText(text);
        Clipboard.SetContent(package);
    }

    /// <summary>
    /// Rows as text for "Copy all", using the Settings delimiter: name, source and value, like the table's columns.
    /// </summary>
    public static string JoinRows(IEnumerable<PropertyEntry> rows)
    {
        var key = App.Settings.Data.CopyAllDelimiter;
        return string.Join(Environment.NewLine, rows.Select(p => Delimiters.Join(key, p.Name, p.Source, p.Value)));
    }

    /// <summary>Right-click (or Shift+F10 / menu key) on a property row: copy its name, value, or both.</summary>
    public static void ShowRowMenu(UIElement sender, ContextRequestedEventArgs args)
    {
        // In a table the row is the element under the pointer or one of its parents (the cell or the TableView
        // row); elsewhere it's the templated element itself, with the entry in its Tag where the template gets
        // no DataContext (the summary facts).
        var entry = RowEntry(args.OriginalSource as DependencyObject, sender)
            ?? (sender as FrameworkElement)?.DataContext as PropertyEntry
            ?? (sender as FrameworkElement)?.Tag as PropertyEntry;
        if (entry == null) return;
        args.Handled = true;

        // Crawled properties carry a property set ID rather than a value.
        bool crawled = entry.Category.Length > 0;
        string valueLabel = crawled ? "property set" : "value";

        var menu = new MenuFlyout();
        menu.Items.Add(CopyItem("Copy name", FluentIcons.Copy, entry.Name));
        menu.Items.Add(CopyItem($"Copy {valueLabel}", FluentIcons.DocumentText, entry.Value, enabled: entry.Value.Length > 0));
        // Plain join (no CSV quoting) so a single name/value pair pastes exactly as shown.
        var pair = $"{entry.Name}{Delimiters.Get(App.Settings.Data.CopyPairDelimiter)}{entry.Value}";
        menu.Items.Add(CopyItem($"Copy name and {valueLabel}", FluentIcons.TextDescription, pair));

        if (args.TryGetPosition(sender, out var point))
            menu.ShowAt(sender, new FlyoutShowOptions { Position = point });
        else
            menu.ShowAt((FrameworkElement)sender);

        static MenuFlyoutItem CopyItem(string text, string glyph, string value, bool enabled = true)
        {
            var item = new MenuFlyoutItem { Text = text, Icon = FluentIcons.Create(glyph), IsEnabled = enabled };
            item.Click += (_, _) => CopyText(value);
            return item;
        }
    }

    /// <summary>The entry of the first element from <paramref name="element"/> up to <paramref name="table"/> that has one.</summary>
    private static PropertyEntry? RowEntry(DependencyObject? element, UIElement table)
    {
        for (; element != null && element != table; element = Microsoft.UI.Xaml.Media.VisualTreeHelper.GetParent(element))
        {
            if (element is FrameworkElement { DataContext: PropertyEntry entry }) return entry;
            if (element is ContentControl { Content: PropertyEntry content }) return content;
        }
        return null;
    }
}

using CommunityToolkit.Mvvm.ComponentModel;
using SharePointSearchExplorer.Helpers;

namespace SharePointSearchExplorer.Models;

/// <summary>
/// A property table's header state: the first column's title and live property count ("37 of 42 properties"),
/// the sort direction of the Property column and the Source column's filter. Used as the columns' Header data
/// with the header templates in PropertyGridResources; the table's columns take a header object and a template, not UI elements.
/// </summary>
public sealed partial class ColumnHeaderInfo(string title) : ObservableObject
{
    public string Title { get; } = title;

    [ObservableProperty]
    private string _count = "";

    /// <summary>Property names Z to A; A to Z otherwise. Clicking the Property header switches it.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(SortGlyph))]
    [NotifyPropertyChangedFor(nameof(SortTip))]
    private bool _sortDescending;

    /// <summary>
    /// The source to show (<see cref="PropertyEntry.ManagedSource"/> or <see cref="PropertyEntry.CrawledSource"/>);
    /// empty for both. Set from the Source header's menu.
    /// </summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsSourceFiltered))]
    private string _sourceFilter = "";

    public string SortGlyph => SortDescending ? FluentIcons.ArrowDown : FluentIcons.ArrowUp;

    public string SortTip => SortDescending ? "Sorted Z to A. Click to sort A to Z." : "Sorted A to Z. Click to sort Z to A.";

    public bool IsSourceFiltered => SourceFilter.Length > 0;

    /// <summary>The header's name for screen readers, which read the Header data as text.</summary>
    public override string ToString() => Title;

    /// <summary>
    /// Fills <paramref name="target"/> with the rows of <paramref name="source"/> (which is sorted by name) that
    /// pass the text filter, the Source filter and the empty toggle, in the header's sort order, and updates
    /// <see cref="Count"/>.
    /// </summary>
    public void Fill(ICollection<PropertyEntry> target, IReadOnlyList<PropertyEntry>? source, string filter, bool showEmpty)
    {
        target.Clear();
        if (source == null)
        {
            Count = "";
            return;
        }
        var rows = source.Where(p =>
            p.Matches(filter)
            && (showEmpty || !p.IsEmpty)
            && (!IsSourceFiltered || p.Source == SourceFilter));
        foreach (var p in SortDescending ? rows.Reverse() : rows) target.Add(p);
        Count = target.Count == source.Count ? $"{source.Count:N0} properties" : $"{target.Count:N0} of {source.Count:N0} properties";
    }
}

namespace SharePointSearchExplorer.Models;

/// <summary>A name/value row shown in the property lists.</summary>
public sealed class PropertyEntry
{
    public const string ManagedSource = "Managed";
    public const string CrawledSource = "Crawled";

    public PropertyEntry(string name, string value, string category = "", string source = "")
    {
        Name = name;
        Value = value;
        Category = category;
        Source = source;
    }

    public string Name { get; }
    public string Value { get; }

    /// <summary>Crawled property set (e.g. "SharePoint", "Office"); empty for managed properties.</summary>
    public string Category { get; }

    /// <summary>
    /// Where a row of the Result Details table comes from (<see cref="ManagedSource"/> or
    /// <see cref="CrawledSource"/>); empty for rows that aren't in that table.
    /// </summary>
    public string Source { get; }

    /// <summary>Crawled rows carry a property set rather than a value, so they never count as empty.</summary>
    public bool IsEmpty => Source != CrawledSource && string.IsNullOrWhiteSpace(Value);

    public bool Matches(string filter) =>
        string.IsNullOrWhiteSpace(filter)
        || Name.Contains(filter, StringComparison.OrdinalIgnoreCase)
        || Value.Contains(filter, StringComparison.OrdinalIgnoreCase)
        || Category.Contains(filter, StringComparison.OrdinalIgnoreCase)
        || Source.Contains(filter, StringComparison.OrdinalIgnoreCase);
}

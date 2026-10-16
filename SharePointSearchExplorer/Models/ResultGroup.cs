using System.Globalization;

namespace SharePointSearchExplorer.Models;

/// <summary>Result Details properties the primary results can be grouped by.</summary>
public enum GroupOption { None, Type, FileType, Author, ModifiedBy, LastModified, Created, Size, Site, ContentClass }

/// <summary>One group of primary results, e.g. every result whose Type is "Word document".</summary>
public sealed class ResultGroup(string name) : List<SearchResultItem>
{
    public string Name { get; } = name;

    public string CountText => Count.ToString("N0", CultureInfo.CurrentCulture);
}

public static class ResultGrouping
{
    private const string NoValue = "(No value)";

    /// <summary>
    /// Groups <paramref name="items"/>, keeping their order inside each group. Dates and sizes fall into
    /// File Explorer-style ranges in their natural order; other values are sorted A–Z with empty values last.
    /// </summary>
    public static List<ResultGroup> Group(IEnumerable<SearchResultItem> items, GroupOption option)
    {
        var today = DateTime.Today;
        var groups = new Dictionary<string, (ResultGroup Group, int Order)>(StringComparer.CurrentCultureIgnoreCase);
        foreach (var item in items)
        {
            var (name, order) = Key(item, option, today);
            if (!groups.TryGetValue(name, out var entry)) groups[name] = entry = (new ResultGroup(name), order);
            entry.Group.Add(item);
        }
        return groups.Values
            .OrderBy(g => g.Order)
            .ThenBy(g => g.Group.Name, StringComparer.CurrentCultureIgnoreCase)
            .Select(g => g.Group)
            .ToList();
    }

    private static (string Name, int Order) Key(SearchResultItem item, GroupOption option, DateTime today) => option switch
    {
        GroupOption.Type => Text(item.TypeLabel),
        GroupOption.FileType => Text(item.FileType),
        GroupOption.Author => Text(item.Author),
        GroupOption.ModifiedBy => Text(item.ModifiedBy),
        GroupOption.LastModified => DateRange(item.LastModified, today),
        GroupOption.Created => DateRange(item.Created, today),
        GroupOption.Size => SizeRange(item.SizeBytes),
        GroupOption.Site => Text(item.Site),
        GroupOption.ContentClass => Text(item.ContentClass),
        _ => Text(""),
    };

    private static (string, int) Text(string value) =>
        string.IsNullOrWhiteSpace(value) ? (NoValue, int.MaxValue) : (value.Trim(), 0);

    private static (string, int) DateRange(DateTime? value, DateTime today)
    {
        if (value is not { } d) return (NoValue, int.MaxValue);
        var date = d.Date;
        int sinceWeekStart = ((int)today.DayOfWeek - (int)CultureInfo.CurrentCulture.DateTimeFormat.FirstDayOfWeek + 7) % 7;
        var weekStart = today.AddDays(-sinceWeekStart);
        var monthStart = new DateTime(today.Year, today.Month, 1);

        if (date > today) return ("Later", 0);
        if (date == today) return ("Today", 1);
        if (date == today.AddDays(-1)) return ("Yesterday", 2);
        if (date >= weekStart) return ("Earlier this week", 3);
        if (date >= weekStart.AddDays(-7)) return ("Last week", 4);
        if (date >= monthStart) return ("Earlier this month", 5);
        if (date >= monthStart.AddMonths(-1)) return ("Last month", 6);
        if (date >= new DateTime(today.Year, 1, 1)) return ("Earlier this year", 7);
        return ("A long time ago", 8);
    }

    private static (string, int) SizeRange(double? bytes)
    {
        const double KB = 1024, MB = KB * 1024, GB = MB * 1024;
        return bytes switch
        {
            null => (NoValue, int.MaxValue),
            < 16 * KB => ("Tiny (0 – 16 KB)", 0),
            < 1 * MB => ("Small (16 KB – 1 MB)", 1),
            < 128 * MB => ("Medium (1 – 128 MB)", 2),
            < 1 * GB => ("Large (128 MB – 1 GB)", 3),
            < 4 * GB => ("Huge (1 – 4 GB)", 4),
            _ => ("Gigantic (> 4 GB)", 5),
        };
    }
}

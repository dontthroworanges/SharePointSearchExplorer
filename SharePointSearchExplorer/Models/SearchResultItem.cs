using System.Globalization;
using System.Text.RegularExpressions;

namespace SharePointSearchExplorer.Models;

/// <summary>A row from PrimaryQueryResult.RelevantResults.</summary>
public sealed partial class SearchResultItem
{
    private readonly Dictionary<string, string> _cells;

    public SearchResultItem(IEnumerable<KeyValuePair<string, string?>> cells)
    {
        _cells = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var ordered = new List<PropertyEntry>();
        foreach (var (key, value) in cells)
        {
            if (string.IsNullOrEmpty(key) || _cells.ContainsKey(key)) continue;
            _cells[key] = value ?? "";
            ordered.Add(new PropertyEntry(key, value ?? ""));
        }
        ReturnedProperties = ordered.OrderBy(p => p.Name, StringComparer.OrdinalIgnoreCase).ToList();

        Path = Get("Path");
        Title = FirstNonEmpty(Get("Title"), Get("Filename"), System.IO.Path.GetFileName(Path.TrimEnd('/')), "(untitled)");
        FileType = FirstNonEmpty(Get("FileType"), Get("FileExtension"), Get("SecondaryFileExtension"));
        WorkId = FirstNonEmpty(Get("WorkId"), Get("DocId"));
        LastModified = ParseDate(Get("LastModifiedTime"));
        Created = ParseDate(Get("Created"));
        Summary = Get("HitHighlightedSummary");
        Author = Get("Author").Replace(";", "; ");
        ModifiedBy = Get("ModifiedBy");
        ContentClass = Get("ContentClass");
        Site = FirstNonEmpty(Get("SiteTitle"), Get("SiteName"), Get("SPWebUrl"));
        SizeBytes = double.TryParse(Get("Size"), NumberStyles.Any, CultureInfo.InvariantCulture, out var size) && size > 0 ? size : null;
        Icon = ResultIcon.For(FileType, ContentClass, IsTrue(Get("IsContainer")), IsTrue(Get("IsDocument")));
        PathUri = AbsoluteUri(Path);
        OpenUri = ViewerUri(Get("SPWebUrl"), Get("UniqueId")) ?? PathUri;
        ParentUri = AbsoluteUri(Get("ParentLink"));

    }

    public string Title { get; }
    public string Path { get; }
    public Uri? PathUri { get; }

    /// <summary>
    /// "Open Result" in a result's menu: the SharePoint viewer, <c>{SPWebUrl}/_layouts/15/viewer.aspx?sourcedoc={UniqueId}</c>.
    /// Falls back to <see cref="PathUri"/> when the result has no SPWebUrl or UniqueId (sites, people, etc.).
    /// </summary>
    public Uri? OpenUri { get; }

    /// <summary>"Open parent location" in a result's right-click menu: the ParentLink managed property.</summary>
    public Uri? ParentUri { get; }
    public string FileType { get; }
    public string WorkId { get; }
    public DateTime? LastModified { get; }
    public DateTime? Created { get; }
    public string Author { get; }
    public string ModifiedBy { get; }
    public string ContentClass { get; }
    public string Site { get; }
    public double? SizeBytes { get; }

    /// <summary>HitHighlightedSummary with &lt;c0&gt; markup intact (rendered by the details view).</summary>
    public string Summary { get; }

    /// <summary>Fluent UI System Icons glyph and type label for the list and details header.</summary>
    public ResultIcon Icon { get; }
    public string Glyph => Icon.Glyph;
    public string TypeLabel => Icon.Label;
    public string? FileTypeIcon => Icon.FileTypeIcon;

    public string LastModifiedText => LastModified is { } d ? $"Last Modified: {d:yyyy-MM-dd}" : "Last Modified: —";

    /// <summary>Every cell returned for this row by the main query, sorted by name.</summary>
    public IReadOnlyList<PropertyEntry> ReturnedProperties { get; }

    // Loaded the first time the item is selected, for the Result Details table.
    internal IReadOnlyList<PropertyEntry>? ManagedProperties { get; set; }
    internal IReadOnlyList<PropertyEntry>? CrawledProperties { get; set; }

    private IReadOnlyList<PropertyEntry>? _allProperties;
    private IReadOnlyList<PropertyEntry>? _allManaged;
    private IReadOnlyList<PropertyEntry>? _allCrawled;

    /// <summary>
    /// The managed and crawled properties in one list sorted by name, each row marked with its source, for
    /// the Result Details table. Null until either list is loaded; rebuilt when either is reloaded.
    /// </summary>
    internal IReadOnlyList<PropertyEntry>? AllProperties
    {
        get
        {
            if (ManagedProperties == null && CrawledProperties == null) return null;
            if (_allProperties == null || _allManaged != ManagedProperties || _allCrawled != CrawledProperties)
            {
                _allManaged = ManagedProperties;
                _allCrawled = CrawledProperties;
                var managed = (ManagedProperties ?? []).Select(p => new PropertyEntry(p.Name, p.Value, source: PropertyEntry.ManagedSource));
                var crawled = (CrawledProperties ?? []).Select(p => new PropertyEntry(p.Name, p.Value, p.Category, PropertyEntry.CrawledSource));
                _allProperties = managed.Concat(crawled).OrderBy(p => p.Name, StringComparer.OrdinalIgnoreCase).ToList();
            }
            return _allProperties;
        }
    }

    public string Get(string key) => _cells.TryGetValue(key, out var v) ? v : "";

    private static Uri? AbsoluteUri(string value) =>
        Uri.TryCreate(value.Trim(), UriKind.Absolute, out var uri) ? uri : null;

    private static Uri? ViewerUri(string webUrl, string uniqueId)
    {
        if (string.IsNullOrWhiteSpace(webUrl) || string.IsNullOrWhiteSpace(uniqueId)) return null;
        return AbsoluteUri($"{webUrl.Trim().TrimEnd('/')}/_layouts/15/viewer.aspx?sourcedoc={Uri.EscapeDataString(uniqueId.Trim())}");
    }

    private static DateTime? ParseDate(string value) =>
        DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var d)
            ? d.ToLocalTime()
            : null;

    private static string FirstNonEmpty(params string[] values) =>
        values.FirstOrDefault(v => !string.IsNullOrWhiteSpace(v)) ?? "";

    private static bool IsTrue(string value) =>
        string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) || value == "1";

    [GeneratedRegex(@"<c\d+>|</c\d+>|<ddd\s*/>", RegexOptions.IgnoreCase)]
    internal static partial Regex HighlightMarkup();
}

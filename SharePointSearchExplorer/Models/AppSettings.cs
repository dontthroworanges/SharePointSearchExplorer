namespace SharePointSearchExplorer.Models;

/// <summary>Everything persisted to settings.json.</summary>
public sealed class AppSettings
{
    public List<ConnectionProfile> Profiles { get; set; } = new();
    public string? CurrentProfileId { get; set; }

    /// <summary>Client ID pre-filled into new profiles.</summary>
    public string DefaultClientId { get; set; } = "";

    /// <summary>Show the Advanced Query Rules and Server Response pages in the navigation rail.</summary>
    public bool ShowAdvancedPages { get; set; }

    public int RowLimit { get; set; } = 50;
    public bool TrimDuplicates { get; set; } = true;

    /// <summary>Delimiter between name and value for right-click "Copy name and value" (a Delimiters key).</summary>
    public string CopyPairDelimiter { get; set; } = "Colon";

    /// <summary>Delimiter between columns for the "Copy all" button (a Delimiters key).</summary>
    public string CopyAllDelimiter { get; set; } = "Comma";

    /// <summary>"Default", "Light" or "Dark".</summary>
    public string Theme { get; set; } = "Default";

    /// <summary>Accent colour as "#RRGGBB" from the Windows colour palette, or empty to use the Windows accent colour.</summary>
    public string AccentColor { get; set; } = "";

    /// <summary>Wrap long text in the results list and property tables (otherwise trim to one line).</summary>
    public bool WordWrap { get; set; } = true;

    /// <summary>Wrap long lines in the Server Response code blocks (otherwise they scroll sideways).</summary>
    public bool WrapCodeBlocks { get; set; }

    /// <summary>Show Microsoft's colourful file type icons in the results list instead of outline icons.</summary>
    public bool UseFileTypeIcons { get; set; }

    /// <summary>Primary results layout: "Comfortable" (icon, title and last modified) or "Dense" (icon and title in shorter rows).</summary>
    public string ResultDensity { get; set; } = "Comfortable";

    /// <summary>Whether property tables start out listing properties with no value.</summary>
    public bool ShowEmptyProperties { get; set; } = true;

    /// <summary>
    /// The Source filter the Result Details table starts with: "Managed", "Crawled" (the PropertyEntry sources),
    /// or empty for both.
    /// </summary>
    public string DefaultPropertySource { get; set; } = "";

    /// <summary>
    /// The managed properties shown in the Result Details summary, in order (at most
    /// <see cref="FavoritePropertyCatalog.MaxCount"/>). Empty when the user removed them all.
    /// </summary>
    public List<string> FavoriteProperties { get; set; } = [.. FavoritePropertyCatalog.Defaults];

    /// <summary>Open the recent queries list when the empty query box is selected.</summary>
    public bool ShowRecentQueriesOnFocus { get; set; } = true;

    /// <summary>Show an X on a recent query in the query box's suggestions, to remove it from the history.</summary>
    public bool AllowQueryBoxDeletion { get; set; } = true;

    /// <summary>Ask for confirmation before the query box's X removes a query.</summary>
    public bool ConfirmQueryBoxDeletion { get; set; } = true;

    /// <summary>Start with the last query (the newest history entry) already in the query box.</summary>
    public bool RestoreLastQuery { get; set; }

    public List<string> History { get; set; } = new();
}

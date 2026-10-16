namespace SharePointSearchExplorer.Models;

/// <summary>Parameters for one call to the SharePoint search endpoint.</summary>
public sealed class SearchQuery
{
    public string QueryText { get; init; } = "*";
    public int StartRow { get; init; }
    public int RowLimit { get; init; } = 50;
    public IReadOnlyList<string> SelectProperties { get; init; } = [];
    public string? Refiners { get; init; }
    public string? RefinementFilters { get; init; }
    public string? HiddenConstraints { get; init; }
    public string? SortList { get; init; }
    public string? QueryTemplate { get; init; }
    public string? SourceId { get; init; }
    public bool TrimDuplicates { get; init; } = true;
}

public sealed record RefinerEntry(string Name, string Token, long Count);

/// <summary>The subset of the search response this tool uses.</summary>
public sealed class SearchResponse
{
    public long TotalRows { get; init; }
    public int ElapsedMilliseconds { get; init; }
    public IReadOnlyList<SearchResultItem> Rows { get; init; } = [];
    public IReadOnlyDictionary<string, IReadOnlyList<RefinerEntry>> Refiners { get; init; } =
        new Dictionary<string, IReadOnlyList<RefinerEntry>>();

    /// <summary>The request and response this was parsed from.</summary>
    public HttpExchange? Http { get; init; }
}

public sealed class SearchException(string message, int? statusCode = null, HttpExchange? http = null) : Exception(message)
{
    public int? StatusCode { get; } = statusCode;

    /// <summary>The failed request (and its response, if one arrived); null when no request was sent.</summary>
    public HttpExchange? Http { get; } = http;
}

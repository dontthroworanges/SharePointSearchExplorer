namespace SharePointSearchExplorer.Models;

/// <summary>One HTTP request to the search endpoint and its response, as shown on the Server Response page.</summary>
public sealed class HttpExchange
{
    public string Method { get; init; } = "";
    public string Url { get; init; } = "";

    /// <summary>Request headers as sent, except that credentials are redacted.</summary>
    public IReadOnlyList<KeyValuePair<string, string>> RequestHeaders { get; init; } = [];
    public string RequestBody { get; init; } = "";

    /// <summary>Null when the request failed before a response arrived (see <see cref="Error"/>).</summary>
    public int? StatusCode { get; init; }
    public string ReasonPhrase { get; init; } = "";
    public string ProtocolVersion { get; init; } = "";
    public IReadOnlyList<KeyValuePair<string, string>> ResponseHeaders { get; init; } = [];
    public string ResponseBody { get; init; } = "";

    /// <summary>Why there is no response (e.g. the site could not be reached).</summary>
    public string? Error { get; init; }

    /// <summary>The query the request was built from (shown on the Debug tab).</summary>
    public SearchQuery? Query { get; init; }

    public bool IsSuccess => StatusCode is >= 200 and < 300;
}
